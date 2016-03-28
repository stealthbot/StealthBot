VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
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
      TabIndex        =   193
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
      TabIndex        =   127
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
      TabIndex        =   126
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
      TabIndex        =   128
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
      TabIndex        =   129
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
      TabIndex        =   130
      Top             =   4920
      Width           =   855
   End
   Begin vbalTreeViewLib6.vbalTreeView tvw 
      Height          =   4620
      Left            =   120
      TabIndex        =   0
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
      Index           =   0
      Left            =   3000
      TabIndex        =   115
      Top             =   0
      Width           =   6615
      Begin VB.CheckBox chkSHR 
         BackColor       =   &H00000000&
         Caption         =   "Shareware"
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
         TabIndex        =   18
         Top             =   3480
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
         TabIndex        =   19
         Top             =   3720
         Width           =   1212
      End
      Begin VB.CheckBox chkSpawn 
         BackColor       =   &H00000000&
         Caption         =   "Use Key as Spawned Client"
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
         TabIndex        =   20
         Top             =   4080
         Width           =   3252
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
         Left            =   240
         MaxLength       =   30
         TabIndex        =   6
         ToolTipText     =   "This account has full control over the bot. Use carefully!"
         Top             =   4080
         Width           =   2532
      End
      Begin VB.CheckBox chkUseRealm 
         BackColor       =   &H00000000&
         Caption         =   "Use Diablo II Realms"
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
         TabIndex        =   9
         Top             =   1575
         Width           =   1935
      End
      Begin VB.ComboBox cboCDKey 
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
         TabIndex        =   3
         Top             =   2280
         Width           =   2535
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
         TabIndex        =   17
         Top             =   3720
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
         TabIndex        =   15
         Top             =   3120
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
         TabIndex        =   10
         Top             =   2280
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
         TabIndex        =   11
         Top             =   2400
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
         TabIndex        =   12
         Top             =   2640
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
         TabIndex        =   13
         Top             =   2760
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
         TabIndex        =   16
         Top             =   3360
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
         TabIndex        =   14
         Top             =   3000
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
         TabIndex        =   1
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
         TabIndex        =   2
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
         TabIndex        =   5
         ToolTipText     =   "This is the channel that the bot will attempt to join when it logs on."
         Top             =   3480
         Width           =   2535
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
         Left            =   3840
         TabIndex        =   8
         Text            =   "."
         Top             =   1560
         Width           =   615
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
         TabIndex        =   7
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
         TabIndex        =   4
         ToolTipText     =   "Only required for Lord of Destruction and The Frozen Throne."
         Top             =   2880
         Width           =   2535
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Bot Owner"
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
         Height          =   252
         Index           =   19
         Left            =   240
         TabIndex        =   194
         Top             =   3840
         Width           =   1212
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
         TabIndex        =   189
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
         TabIndex        =   188
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label1 
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
         TabIndex        =   140
         Top             =   2160
         Width           =   735
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         Index           =   2
         X1              =   4680
         X2              =   4800
         Y1              =   3120
         Y2              =   3240
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   4680
         X2              =   4800
         Y1              =   2760
         Y2              =   2880
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   4680
         X2              =   4800
         Y1              =   2400
         Y2              =   2520
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   3000
         X2              =   3000
         Y1              =   960
         Y2              =   4080
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Product"
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
         TabIndex        =   139
         Top             =   2040
         Width           =   615
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   360
         X2              =   6240
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Trigger"
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
         TabIndex        =   137
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Server"
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
         TabIndex        =   136
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Username"
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
         Left            =   240
         TabIndex        =   135
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Password"
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
         TabIndex        =   134
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "CDKey"
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
         TabIndex        =   133
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Home Channel"
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
         TabIndex        =   132
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Expansion CDKey"
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
         TabIndex        =   131
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label1 
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
         Index           =   7
         Left            =   360
         TabIndex        =   138
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
      TabIndex        =   117
      Top             =   0
      Width           =   6615
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
         Left            =   2040
         TabIndex        =   31
         Text            =   "10000"
         Top             =   4080
         Width           =   735
      End
      Begin VB.CheckBox chkNoColoring 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Disable channel list name coloring"
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
         Left            =   3120
         TabIndex        =   29
         ToolTipText     =   "Name coloring changes the color of people's names in the channel list based on their status or activity."
         Top             =   2400
         Width           =   3015
      End
      Begin VB.CheckBox chkNoAutocomplete 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Disable name autocompletion"
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
         TabIndex        =   25
         ToolTipText     =   "Checking this box prevents the highlighted display of suggested usernames as you type in the send box."
         Top             =   2160
         Width           =   2535
      End
      Begin VB.CheckBox chkNoTray 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Do not minimize to the System Tray"
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
         Left            =   3120
         TabIndex        =   28
         ToolTipText     =   "Disable minimization to the System Tray (only to the Taskbar)."
         Top             =   2040
         Width           =   3015
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
         TabIndex        =   24
         ToolTipText     =   "Flash the bot window when events occur."
         Top             =   1800
         Width           =   2535
      End
      Begin VB.CheckBox chkUTF8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Use UTF-8 encoding/decoding when processing and sending messages"
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
         Left            =   3120
         TabIndex        =   27
         ToolTipText     =   "Blizzard games encode their messages to UTF-8 format. Enable this setting to properly see special characters sent by these games."
         Top             =   1440
         Value           =   1  'Checked
         Width           =   3015
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
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   960
         Width           =   3135
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
         TabIndex        =   23
         ToolTipText     =   "Enable/disable the splash screen on startup."
         Top             =   1440
         Width           =   2535
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
         Left            =   360
         TabIndex        =   22
         ToolTipText     =   "Enable/disable chat filtering (lowers CPU usage)"
         Top             =   1080
         Width           =   2535
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
         TabIndex        =   21
         ToolTipText     =   "Enable/disable Join and Leave messages"
         Top             =   720
         Width           =   2535
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
         TabIndex        =   30
         Top             =   3000
         Width           =   5895
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
         Left            =   2040
         TabIndex        =   32
         Text            =   "0"
         Top             =   4410
         Width           =   735
      End
      Begin VB.Label lblBacklog 
         BackColor       =   &H00000000&
         Caption         =   "Maximum backlog size"
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
         TabIndex        =   197
         Top             =   4110
         Width           =   1575
      End
      Begin VB.Label lblBacklogSize 
         BackColor       =   &H00000000&
         Caption         =   "  bytes (set to 0 for unlimited)"
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
         Left            =   2790
         TabIndex        =   196
         Top             =   4110
         Width           =   3015
      End
      Begin VB.Label Label8 
         BackColor       =   &H00000000&
         Caption         =   "Timestamp Settings"
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
         Left            =   3120
         TabIndex        =   158
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lbl5 
         BackColor       =   &H00000000&
         Caption         =   "  megabytes (set to 0 for unlimited)"
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
         Left            =   2790
         TabIndex        =   150
         Top             =   4440
         Width           =   3015
      End
      Begin VB.Label lbl5 
         BackColor       =   &H00000000&
         Caption         =   "Maximum logfile size"
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
         Left            =   360
         TabIndex        =   149
         Top             =   4440
         Width           =   1575
      End
      Begin VB.Label lbl5 
         BackColor       =   &H00000000&
         Caption         =   "Channel text logging"
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
         Left            =   360
         TabIndex        =   148
         Top             =   2760
         Width           =   2535
      End
      Begin VB.Label lbl5 
         BackColor       =   &H00000000&
         Caption         =   $"frmSettings.frx":0000
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
         Index           =   5
         Left            =   360
         TabIndex        =   147
         Top             =   3360
         Width           =   5895
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   3
         X1              =   360
         X2              =   6240
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label1 
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
         Index           =   12
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
      Index           =   3
      Left            =   3000
      TabIndex        =   118
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
         TabIndex        =   39
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
         TabIndex        =   38
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
         TabIndex        =   40
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
         TabIndex        =   41
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
         TabIndex        =   42
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
         TabIndex        =   43
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
         TabIndex        =   44
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
         TabIndex        =   47
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
         TabIndex        =   48
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
         TabIndex        =   37
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
         TabIndex        =   49
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
         TabIndex        =   45
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
         TabIndex        =   46
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
         TabIndex        =   35
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
         TabIndex        =   33
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox txtChanSize 
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
         Left            =   4560
         MaxLength       =   2
         TabIndex        =   36
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtChatSize 
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
         Left            =   4560
         MaxLength       =   2
         TabIndex        =   34
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
         TabIndex        =   167
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label lblEg 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3360
         TabIndex        =   166
         Top             =   2520
         Width           =   2895
      End
      Begin VB.Label Label1 
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
         Index           =   14
         Left            =   360
         TabIndex        =   165
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label3 
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
         Left            =   360
         TabIndex        =   164
         Top             =   2280
         Width           =   4335
      End
      Begin VB.Label Label4 
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
         Left            =   360
         TabIndex        =   163
         Top             =   3360
         Width           =   135
      End
      Begin VB.Label Label5 
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
         Left            =   1320
         TabIndex        =   162
         Top             =   3360
         Width           =   135
      End
      Begin VB.Label Label6 
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
         Left            =   2280
         TabIndex        =   161
         Top             =   3360
         Width           =   255
      End
      Begin VB.Label Label7 
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
         Left            =   3360
         TabIndex        =   160
         Top             =   3000
         Width           =   2415
      End
      Begin VB.Label Label8 
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
         Index           =   5
         Left            =   3360
         TabIndex        =   159
         Top             =   3240
         Width           =   255
      End
      Begin VB.Label Label8 
         BackColor       =   &H00000000&
         Caption         =   "Size"
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
         Left            =   4200
         TabIndex        =   157
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label8 
         BackColor       =   &H00000000&
         Caption         =   "Channel List"
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
         TabIndex        =   156
         ToolTipText     =   "Changes the font settings for the channel list."
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label8 
         BackColor       =   &H00000000&
         Caption         =   "Chat Window"
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
         Left            =   360
         TabIndex        =   155
         ToolTipText     =   "Changes the font setting for the main chat window."
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label8 
         BackColor       =   &H00000000&
         Caption         =   "Font"
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
         TabIndex        =   154
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label8 
         BackColor       =   &H00000000&
         Caption         =   "Size"
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
         Left            =   4200
         TabIndex        =   153
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label8 
         BackColor       =   &H00000000&
         Caption         =   "Font"
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
         TabIndex        =   152
         Top             =   1200
         Width           =   495
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   4
         X1              =   360
         X2              =   6240
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label1 
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
         Index           =   13
         Left            =   360
         TabIndex        =   151
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
         TabIndex        =   192
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
      TabIndex        =   119
      Top             =   0
      Width           =   6615
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
         TabIndex        =   99
         ToolTipText     =   "Ban Evasion attempts to keep people who are banned out of your channel."
         Top             =   2880
         Width           =   2295
      End
      Begin VB.CheckBox chkIdleKick 
         BackColor       =   &H00000000&
         Caption         =   "Kick instead of ban"
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
         Left            =   4440
         TabIndex        =   102
         ToolTipText     =   "Instead of banning idle users, the bot will simply kick them."
         Top             =   720
         Width           =   1815
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
         Left            =   3000
         TabIndex        =   104
         ToolTipText     =   "Ban Warcraft III users who have the Peon icon."
         Top             =   1440
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
         Left            =   3000
         MaxLength       =   180
         TabIndex        =   114
         Text            =   "You are below the required level for entry."
         Top             =   4180
         Width           =   3375
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
         TabIndex        =   111
         Top             =   2520
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
         TabIndex        =   110
         Top             =   2280
         Width           =   1812
      End
      Begin VB.CheckBox chkPhrasebans 
         BackColor       =   &H00000000&
         Caption         =   "Enable Phrasebanning"
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
         ToolTipText     =   "Ban unsafelisted users who state banned phrases."
         Top             =   720
         Width           =   1935
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
         TabIndex        =   94
         ToolTipText     =   "Ban squelched users."
         Top             =   1080
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
         Left            =   3000
         TabIndex        =   105
         Top             =   2040
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
         TabIndex        =   109
         Top             =   2040
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
         Left            =   3000
         TabIndex        =   106
         Top             =   2280
         Width           =   975
      End
      Begin VB.CheckBox chkCBan 
         BackColor       =   &H00000000&
         Caption         =   "Warcraft II BNE"
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
         TabIndex        =   108
         Top             =   2760
         Width           =   1455
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
         Left            =   3000
         TabIndex        =   107
         Top             =   2520
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
         TabIndex        =   95
         ToolTipText     =   "Ban unsafelisted users that talk."
         Top             =   1440
         Width           =   1935
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
         TabIndex        =   100
         Text            =   "Lockdown Enabled"
         Top             =   4180
         Width           =   2535
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
         TabIndex        =   98
         ToolTipText     =   "Ban unsafelisted users who join the channel."
         Top             =   2520
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
         TabIndex        =   96
         ToolTipText     =   "Kick users who yell (uppercase message longer than 5 letters)"
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox txtBanW3 
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
         Left            =   5400
         MaxLength       =   25
         TabIndex        =   113
         Top             =   3480
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
         TabIndex        =   97
         ToolTipText     =   "Ban users with a UDP plug."
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox txtBanD2 
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
         Left            =   3840
         MaxLength       =   25
         TabIndex        =   112
         Top             =   3480
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
         Left            =   3000
         TabIndex        =   101
         ToolTipText     =   "Ban users who have been idle for X seconds."
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtIdleBanDelay 
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
         MaxLength       =   25
         TabIndex        =   103
         Top             =   1080
         Width           =   615
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
         Index           =   0
         Left            =   3000
         TabIndex        =   184
         Top             =   3940
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
         Index           =   2
         Left            =   3000
         TabIndex        =   174
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
         Height          =   615
         Index           =   3
         Left            =   3000
         TabIndex        =   173
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lblMod 
         BackColor       =   &H00000000&
         Caption         =   "Protection ban message"
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
         TabIndex        =   172
         ToolTipText     =   "Shorter is better"
         Top             =   3940
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
         Index           =   1
         Left            =   3000
         TabIndex        =   171
         Top             =   3240
         Width           =   2295
      End
      Begin VB.Label lblMod 
         Alignment       =   1  'Right Justify
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
         Index           =   4
         Left            =   3000
         TabIndex        =   170
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label lblMod 
         Alignment       =   1  'Right Justify
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
         Index           =   6
         Left            =   4320
         TabIndex        =   169
         Top             =   3480
         Width           =   975
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   5
         X1              =   360
         X2              =   6240
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label1 
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
         Index           =   15
         Left            =   360
         TabIndex        =   168
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
      TabIndex        =   121
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
         TabIndex        =   85
         Top             =   840
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
         TabIndex        =   88
         Top             =   1920
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
         TabIndex        =   86
         Top             =   1200
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
         TabIndex        =   87
         Top             =   1560
         Width           =   975
      End
      Begin VB.CheckBox chkIdles 
         BackColor       =   &H00000000&
         Caption         =   "Show anti-idle messages"
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
         TabIndex        =   83
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
         TabIndex        =   89
         Text            =   "/me is a %ver"
         Top             =   1800
         Width           =   4335
      End
      Begin VB.TextBox txtIdleWait 
         Alignment       =   2  'Center
         BackColor       =   &H00993300&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3000
         TabIndex        =   84
         Text            =   "6"
         Top             =   1200
         Width           =   495
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         Index           =   4
         X1              =   3840
         X2              =   5160
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         Index           =   3
         X1              =   5040
         X2              =   5280
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         Index           =   2
         X1              =   5040
         X2              =   5280
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   5040
         X2              =   5280
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   5040
         X2              =   5040
         Y1              =   1080
         Y2              =   2040
      End
      Begin VB.Label lblIdle 
         BackColor       =   &H00000000&
         Caption         =   "Idle message type"
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
         Left            =   3840
         TabIndex        =   182
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lblIdle 
         BackColor       =   &H00000000&
         Caption         =   "Delay between messages (minutes)"
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
         TabIndex        =   181
         Top             =   1200
         Width           =   3015
      End
      Begin VB.Label lblIdle 
         BackColor       =   &H00000000&
         Caption         =   "Idle message"
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
         TabIndex        =   180
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   7
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
         TabIndex        =   178
         Top             =   2280
         Width           =   5895
      End
      Begin VB.Label Label1 
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
         Index           =   17
         Left            =   360
         TabIndex        =   179
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
      TabIndex        =   122
      Top             =   0
      Width           =   6615
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
         Left            =   360
         TabIndex        =   60
         Top             =   4440
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CheckBox chkD2Naming 
         BackColor       =   &H00000000&
         Caption         =   "Use Diablo II naming conventions"
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
         TabIndex        =   70
         ToolTipText     =   "Show usernames with Diablo II naming conventions."
         Top             =   4080
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
         TabIndex        =   69
         ToolTipText     =   "Show usernames with all gateways."
         Top             =   3720
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
         TabIndex        =   68
         ToolTipText     =   "Show usernames as they would appear to a WarCraft III user."
         Top             =   3720
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
         TabIndex        =   67
         ToolTipText     =   "Show usernames as they would appear to a StarCraft or WarCraft II user."
         Top             =   3360
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
         TabIndex        =   66
         ToolTipText     =   "Show usernames as they would appear to the selected client."
         Top             =   3360
         Value           =   -1  'True
         Width           =   1215
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
         Height          =   375
         Left            =   360
         TabIndex        =   59
         Top             =   4035
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CheckBox chkDisableSuffix 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Disable suffix box"
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
         TabIndex        =   55
         ToolTipText     =   "Disables the smaller suffix box to the right of the box you type in to send text to Battle.net"
         Top             =   2640
         Width           =   2535
      End
      Begin VB.CheckBox chkDisablePrefix 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Disable prefix box"
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
         TabIndex        =   54
         ToolTipText     =   "Disables the smaller prefix box to the left of the box you type in to send text to Battle.net"
         Top             =   2280
         Width           =   2535
      End
      Begin VB.CheckBox chkLogAllCommands 
         BackColor       =   &H00000000&
         Caption         =   "Log all commands"
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
         TabIndex        =   63
         ToolTipText     =   "Any commands issued to the bot will be logged in a file in the bot's Logs folder called 'commandlog.txt'."
         Top             =   1890
         Width           =   2295
      End
      Begin VB.CheckBox chkLogDBActions 
         BackColor       =   &H00000000&
         Caption         =   "Log database changes"
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
         TabIndex        =   62
         ToolTipText     =   "Any actions that change the database will be logged in the log folder in a file called 'database.txt'."
         Top             =   1560
         Width           =   2295
      End
      Begin VB.CheckBox chkShowOffline 
         Alignment       =   1  'Right Justify
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
         Left            =   360
         TabIndex        =   57
         ToolTipText     =   "Determines whether or not offline friends are hidden from /f list"
         Top             =   3360
         Width           =   2535
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
         Left            =   360
         TabIndex        =   56
         ToolTipText     =   "Enables automatic URL detection and highlighting in the main chat window."
         Top             =   3000
         Width           =   2535
      End
      Begin VB.CheckBox chkDoNotUsePacketFList 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Do not use 0x65 internal friends' list"
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
         TabIndex        =   58
         ToolTipText     =   "Disable the internal friends' list (alternative channel list)"
         Top             =   3675
         Width           =   2535
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
         TabIndex        =   65
         Top             =   2775
         Width           =   2535
      End
      Begin VB.CheckBox chkMinimizeOnStartup 
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
         Left            =   3480
         TabIndex        =   61
         ToolTipText     =   "Automatically minimize on startup."
         Top             =   1200
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
         TabIndex        =   50
         ToolTipText     =   "Allow commands such as .next and .back that change your current Winamp song."
         Top             =   840
         Width           =   2535
      End
      Begin VB.CheckBox chkWhisperCmds 
         Alignment       =   1  'Right Justify
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
         Left            =   360
         TabIndex        =   52
         ToolTipText     =   "Whisper the return messages of all bot commands."
         Top             =   1560
         Width           =   2535
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
         TabIndex        =   51
         ToolTipText     =   "Enable/disable ProfileAmp - writes Winamp's currently played song to your profile every 30 seconds"
         Top             =   1200
         Width           =   2535
      End
      Begin VB.CheckBox chkMail 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Check users' mail"
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
         TabIndex        =   53
         ToolTipText     =   "Enable/disable checking of the mail.ini file when people join."
         Top             =   1920
         Width           =   2535
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
         TabIndex        =   64
         ToolTipText     =   "The bot will join a specified channel when kicked, instead of rejoining."
         Top             =   2160
         Width           =   2895
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gateway naming convention:"
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
         Index           =   6
         Left            =   3480
         TabIndex        =   198
         Top             =   3120
         Width           =   2100
      End
      Begin VB.Label Label8 
         BackColor       =   &H00000000&
         Caption         =   "Backup channel:"
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
         Left            =   3480
         TabIndex        =   187
         Top             =   2550
         Width           =   2415
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   9
         X1              =   3120
         X2              =   3120
         Y1              =   840
         Y2              =   4560
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   8
         X1              =   360
         X2              =   6240
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label1 
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
         Index           =   18
         Left            =   360
         TabIndex        =   183
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
      TabIndex        =   116
      Top             =   0
      Width           =   6615
      Begin VB.CheckBox chkConnectOnStartup 
         BackColor       =   &H00000000&
         Caption         =   "Connect on startup"
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
         TabIndex        =   73
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
         TabIndex        =   75
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
         ItemData        =   "frmSettings.frx":00D6
         Left            =   2520
         List            =   "frmSettings.frx":00D8
         TabIndex        =   72
         Text            =   "cboBNLSServer"
         Top             =   1200
         Width           =   3735
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
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   74
         Text            =   "1000"
         Top             =   3000
         Width           =   615
      End
      Begin VB.OptionButton optSocks5 
         BackColor       =   &H00000000&
         Caption         =   "SOCKS5"
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
         TabIndex        =   82
         Top             =   3480
         Width           =   735
      End
      Begin VB.OptionButton optSocks4 
         BackColor       =   &H00000000&
         Caption         =   "SOCKS4"
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
         TabIndex        =   81
         Top             =   3480
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
         Left            =   3960
         MaxLength       =   5
         TabIndex        =   80
         Top             =   4200
         Width           =   1095
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
         Left            =   3960
         MaxLength       =   15
         TabIndex        =   79
         Top             =   3840
         Width           =   2175
      End
      Begin VB.CheckBox chkUseProxies 
         BackColor       =   &H00000000&
         Caption         =   "Use a proxy"
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
         TabIndex        =   78
         ToolTipText     =   "Routes your Battle.net and/or BNLS connection through a SOCKS4 or 5 proxy."
         Top             =   3480
         Width           =   1215
      End
      Begin VB.CheckBox chkUDP 
         BackColor       =   &H00000000&
         Caption         =   "Use Lag Plug"
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
         TabIndex        =   77
         ToolTipText     =   "Sets whether or not you have a lag plug when you sign on. If you don't know what this is, leave it off."
         Top             =   4200
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
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   76
         Top             =   3840
         Width           =   2175
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
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   71
         Top             =   840
         Width           =   4335
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00FFFFFF&
         X1              =   2640
         X2              =   240
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Label lbl5 
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
         Index           =   13
         Left            =   4320
         TabIndex        =   200
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label lbl5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Email Address"
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
         Left            =   3000
         TabIndex        =   199
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FFFFFF&
         X1              =   2880
         X2              =   6360
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Label lbl5 
         BackColor       =   &H00000000&
         Caption         =   "BNLS server, if applicable:"
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
         Index           =   12
         Left            =   360
         TabIndex        =   195
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label lbl5 
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
         Index           =   11
         Left            =   2400
         TabIndex        =   191
         Top             =   3000
         Width           =   255
      End
      Begin VB.Label lbl5 
         BackColor       =   &H00000000&
         Caption         =   "Reconnect delay"
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
         Left            =   360
         TabIndex        =   190
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label lbl5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Port"
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
         TabIndex        =   186
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label lbl5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "IP address"
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
         TabIndex        =   185
         Top             =   3840
         Width           =   855
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         X1              =   2760
         X2              =   2760
         Y1              =   2520
         Y2              =   4680
      End
      Begin VB.Label lbl5 
         BackColor       =   &H00000000&
         Caption         =   "Ping spoofing"
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
         TabIndex        =   146
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
         TabIndex        =   145
         Top             =   1920
         Width           =   6015
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbl5 
         BackColor       =   &H00000000&
         Caption         =   "Connection method:"
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
         TabIndex        =   144
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label lbl5 
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
         Index           =   3
         Left            =   360
         TabIndex        =   143
         Top             =   1680
         Width           =   5895
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   2
         X1              =   360
         X2              =   6240
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label1 
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
         Index           =   11
         Left            =   360
         TabIndex        =   141
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
      TabIndex        =   120
      Top             =   0
      Width           =   6615
      Begin VB.CheckBox chkWhisperGreet 
         BackColor       =   &H00000000&
         Caption         =   "Whisper the greet message"
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
         TabIndex        =   91
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
         TabIndex        =   92
         Text            =   "Welcome to %c, %0! I am a %v."
         Top             =   1800
         Width           =   5895
      End
      Begin VB.CheckBox chkGreetMsg 
         BackColor       =   &H00000000&
         Caption         =   "Greet users who join the channel"
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
         TabIndex        =   90
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
         TabIndex        =   177
         Top             =   2280
         Width           =   5892
      End
      Begin VB.Label lblIdle 
         BackColor       =   &H00000000&
         Caption         =   "Greet Message"
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
         TabIndex        =   176
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   6
         X1              =   360
         X2              =   6240
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label1 
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
         Index           =   16
         Left            =   360
         TabIndex        =   175
         Top             =   240
         Width           =   4815
      End
   End
   Begin VB.Frame fraPanel 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   4815
      Index           =   8
      Left            =   3000
      TabIndex        =   123
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
         TabIndex        =   125
         Top             =   960
         Width           =   6015
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   360
         X2              =   6240
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Welcome to &StealthBot"
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
         TabIndex        =   124
         Top             =   240
         Width           =   3255
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
    
    '##########################################
    ' TREEVIEW INITIALIZATION CODE
    '##########################################
    
    lMouseOver = vbWhite
    
    With tvw
        
        .nodes.Clear
        
        Set nRoot = .nodes.Add(, etvwFirst, "root", "StealthBot Settings")
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
    End With
    
    With cboConnMethod
        .AddItem "BNLS - Battle.net Logon Server"
        .AddItem "ADVANCED - Local hashing"
        .ListIndex = 0
    End With
    
    With cboBNLSServer
        .AddItem "Automatic (Server Finder)"
        
        ' If the user has a server set, add it.
        If Len(Config.BnlsServer) > 0 Then
            AddBNLSServer Config.BnlsServer
        End If
        
        ' Figure out the default selection
        If Config.UseBnlsFinder Or Len(Config.BnlsServer) = 0 Then
            .ListIndex = 0
        ElseIf Len(Config.BnlsServer) > 0 Then
            .ListIndex = GetBnlsIndex(Config.BnlsServer)
        Else
            .ListIndex = 0
        End If
        
        ' Add servers from the user's local list.
        If LenB(Dir$(GetFilePath(FILE_BNLS_LIST))) > 0 Then
            With cboBNLSServer
                i = FreeFile
                
                Open GetFilePath(FILE_BNLS_LIST) For Input As #i
                    While Not EOF(i)
                        Line Input #i, s
                        
                        If Len(s) > 0 Then AddBNLSServer s
                    Wend
                Close #i
            End With
        End If
    End With
    
    With cboLogging
        .AddItem "Full logging - text is logged and a dated logfile is saved during operation."
        .AddItem "Partial logging - text is logged. The logfile is deleted on shutdown."
        .AddItem "No logging."
        .ListIndex = 0
    End With
    
    With cboTimestamp
        .AddItem "[HH:MM:SS PM] - Seconds with time of day"
        .AddItem "[HH:MM:SS] - Seconds"
        .AddItem "[HH:MM:SS:mmm] - Milliseconds"
        .AddItem "No timestamp"
        .ListIndex = 0
    End With
    
    lblGreetVars.Caption = "Greet Message Variables: (Suggest more! email stealth@stealthbot.net) " & vbNewLine & _
        "%c = Current channel" & vbNewLine & _
        "%0 = Username of the person who joins" & vbNewLine & _
        "%1 = Bot's current username" & vbNewLine & _
        "%p = Ping of user who joins" & vbNewLine & _
        "%v = The bot's current version" & vbNewLine & _
        "%a = Database access of the person who joins" & vbNewLine & _
        "%f = Database flags of the person who joins" & vbNewLine & _
        "%t = Current time" & vbNewLine & _
        "%d = Current date"

    lblIdleVars.Caption = "Idle message variables: (Suggest more! email stealth@stealthbot.net) " & vbNewLine & _
        "%c = Current channel" & vbNewLine & _
        "%me = Current username" & vbNewLine & _
        "%v = Bot version" & vbNewLine & _
        "%botup = Bot uptime" & vbNewLine & _
        "%cpuup = System uptime" & vbNewLine & _
        "%mp3 = Current MP3" & vbNewLine & _
        "%quote = Random quote" & vbNewLine & _
        "%rnd = Random person in the channel" & vbNewLine
    
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
        ShowPanel spSplash, 1
    Else
        ShowPanel CInt(Config.LastSettingsPanel)
    End If
End Sub

Private Sub Form_GotFocus()
    If Len(cboCDKey.Text) = 0 And cboCDKey.ListCount > 0 Then
        cboCDKey.ListIndex = 1
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

Private Sub SetHashPath(ByVal path As String)
    lblHashPath.Caption = path
    lblHashPath.ToolTipText = path
    
    ' Try to adjust font size for long paths
    If ((Len(path) > 80) And (Not InStr(path, Space(1)) > 0)) Then
        If Len(path) > 100 Then
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
        Case "op_greets":       KeyToIndex = 5
        Case "op_idles":        KeyToIndex = 6
        Case "op_misc":         KeyToIndex = 7
        
        Case Else:              KeyToIndex = 8
    End Select
    
End Function

Sub ShowPanel(ByVal index As enuSettingsPanels, Optional ByVal Mode As Byte = 0)

    Static ActivePanel As Integer
    
    If PanelsInitialized Then
        Dim nod As cTreeViewNode
        Dim i As Integer
        If index <> 8 Then
            For i = 1 To tvw.NodeCount
                Set nod = tvw.nodes.Item(i)
                If Not nod.Selected And KeyToIndex(nod.key) = index Then
                    nod.Selected = True
                    Exit For
                End If
            Next i
            Set nod = Nothing
        End If
        If Mode = 1 Then
            fraPanel(KeyToIndex("splash")).ZOrder vbBringToFront
            ActivePanel = KeyToIndex("splash")
        Else
            'fraPanel(ActivePanel).ZOrder vbSendToBack
            fraPanel(index).ZOrder vbBringToFront
            ActivePanel = index
            Config.LastSettingsPanel = ActivePanel
            
            'Debug.Print "Writing: " & ActivePanel
        End If
    End If
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
    frmCustomInputBox.Show
End Sub

Private Sub cmdWebsite_Click()
    Call frmChat.mnuHelpWebsite_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call frmChat.DeconstructSettings
    Set colProfiles = Nothing
End Sub

Sub lblAddCurrentKey_Click()
    Dim i As Integer
    Dim s As String
    
    s = CDKeyReplacements(cboCDKey.Text)
    
    If cboCDKey.ListCount > -1 Then
        For i = 0 To cboCDKey.ListCount
            If StrComp(cboCDKey.List(i), s, vbTextCompare) = 0 Then
                Exit Sub
            End If
        Next i
    End If
    
    cboCDKey.AddItem s
End Sub

Private Sub lblManageKeys_Click()
    If LenB(cboCDKey.Text) > 0 Then
        Call lblAddCurrentKey_Click
    End If
    
    Call WriteCDKeys(cboCDKey)
    frmManageKeys.Show
End Sub

Sub lblAddCurrentKey_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblAddCurrentKey.ForeColor = vbBlue
    lblManageKeys.ForeColor = vbWhite
End Sub

Sub lblManageKeys_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblAddCurrentKey.ForeColor = vbWhite
    lblManageKeys.ForeColor = vbBlue
End Sub

Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblAddCurrentKey.ForeColor = vbWhite
    lblManageKeys.ForeColor = vbWhite
End Sub

Sub fraPanel_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
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
'Private Sub tvw_NodeClick(node As vbalTreeViewLib6.cTreeViewNode)
'    Call tvw_SelectedNodeChanged
'End Sub

Private Sub tvw_SelectedNodeChanged()
    If FirstRun = 0 Then
        ShowPanel KeyToIndex(tvw.SelectedItem.key)
    Else
        FirstRun = 0
    End If
End Sub

Private Function DoCDKeyLengthCheck(ByVal sKey As String, ByVal sProd As String) As Boolean
    sKey = CDKeyReplacements(sKey)
    
    DoCDKeyLengthCheck = True
    If Config.SetKeyIgnoreLength Then Exit Function
    
    Select Case sProd
        Case "STAR", "SEXP"
            If ((Len(sKey) <> 13) And (Len(sKey) <> 26)) Then DoCDKeyLengthCheck = False
            
        Case "D2DV", "D2XP"
            If ((Len(sKey) <> 16) And (Len(sKey) <> 26)) Then DoCDKeyLengthCheck = False
            
        Case "W2BN"
            If (Len(sKey) <> 16) Then DoCDKeyLengthCheck = False
            
        Case "WAR3", "W3XP"
            If (Len(sKey) <> 26) Then DoCDKeyLengthCheck = False
        
        Case "SSHR", "DSHR", "DRTL"
            DoCDKeyLengthCheck = True
        
        Case "JSTR"
            If (Len(sKey) <> 13) Then DoCDKeyLengthCheck = False
        
        Case Else
            DoCDKeyLengthCheck = False
    End Select
End Function

Private Function SaveSettings() As Boolean
    Dim s As String
    Dim Clients(6) As String
    Dim i As Long, j As Long
    
    '// First, CDKey Length check and corresponding stuff that needs to run first:
    Select Case True
        Case optSTAR.value:
            If (chkSHR.value) Then
                s = "SSHR"
            ElseIf (chkJPN.value) Then
                s = "JSTR"
            Else
                s = "STAR"
            End If
        Case optSEXP.value: s = "SEXP"
        Case optD2DV.value: s = "D2DV"
        Case optD2XP.value: s = "D2XP"
        Case optWAR3.value: s = "WAR3"
        Case optW3XP.value: s = "W3XP"
        Case optW2BN.value: s = "W2BN"
        Case optDRTL.value:
            If (chkSHR.value) Then
                s = "DSHR"
            Else
                s = "DRTL"
            End If
        'Case optCHAT.Value: s = "CHAT"
    End Select
    
    If Not DoCDKeyLengthCheck(cboCDKey.Text, s) Then
        If MsgBox("Your CD key is of an invalid Length for the product you have chosen. Do you want to save anyway?", vbExclamation + vbYesNo, "StealthBot Settings") = vbNo Then
            ShowPanel spConnectionConfig
            cboCDKey.SetFocus
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
    
    Config.Product = StrReverse(s)
    
    '// The rest of the basic config now
    Config.Username = txtUsername.Text
    Config.Password = txtPassword.Text
    Config.CdKey = CDKeyReplacements(cboCDKey.Text)
    Config.ExpKey = CDKeyReplacements(txtExpKey.Text)
    Config.HomeChannel = txtHomeChan.Text
    Config.Server = cboServer.Text
    Config.UseSpawnKey = CBool(chkSpawn.value)
    Config.Trigger = txtTrigger.Text
    Config.BotOwner = txtOwner.Text
    Config.UseRealm = CBool(chkUseRealm.value)
    
    ' Advanced connection settings
    Config.UseBnls = CBool(cboConnMethod.ListIndex = 0)
    Config.UseBnlsFinder = CBool(cboBNLSServer.ListIndex = 0)
    
    If cboBNLSServer.ListIndex > 0 Then
        Config.BnlsServer = cboBNLSServer.Text
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
            i = FreeFile
            
            ' Save the list of servers to a file
            Open GetFilePath(FILE_BNLS_LIST) For Output As #i
                For j = 1 To .ListCount
                    If Not (Len(.List(j)) = 0) Then Print #i, .List(j)
                Next j
            Close #i
        End If
    End With
    
    Config.ConnectOnStartup = CBool(chkConnectOnStartup.value)
    Config.RegisterEmailDefault = Trim$(txtEmail.Text)
    Config.PingSpoofing = cboSpoof.ListIndex
    Config.UseProxy = CBool(chkUseProxies.value)
    Config.ProxyPort = CLng(Trim$(txtProxyPort.Text))
    Config.ProxyIP = Trim$(txtProxyIP.Text)
    Config.ProxyIsSocks5 = CBool(optSocks5.value)
    Config.UseUDP = CBool(chkUDP.value)

    Config.ReconnectDelay = CLng(txtReconDelay.Text)
    
    '// General Interface settings
    Config.ShowJoinLeaves = CBool(chkJoinLeaves.value)
    Config.UseChatFilters = CBool(chkFilter.value)
    Config.ShowSplashScreen = CBool(chkSplash.value)
    Config.FlashOnEvents = CBool(chkFlash.value)
    Config.Timestamp = cboTimestamp.ListIndex
    Config.UTF8 = CBool(chkUTF8.value)
    
    If (cboLogging.ListIndex = 0) Then
        Config.LoggingLevel = 2
    ElseIf (cboLogging.ListIndex = 1) Then
        Config.LoggingLevel = 1
    ElseIf (cboLogging.ListIndex = 2) Then
        Config.LoggingLevel = 0
    End If
    
    Config.MaxBacklogSize = CLng(txtMaxBackLogSize.Text)
    Config.MaxLogFileSize = CLng(txtMaxLogSize.Text)
    Config.MinimizeToTray = Not CBool(chkNoTray.value)
    Config.DisableAutoComplete = CBool(chkNoAutocomplete.value)
    Config.DisableNameColoring = CBool(chkNoColoring.value)
    
    '// Misc General Settings
    Config.UseProfileAmp = CBool(chkPAmp.value)
    Config.WhisperResponses = CBool(chkWhisperCmds.value)
    Config.EnableMail = CBool(chkMail.value)
    Config.DisablePrefixBox = CBool(chkDisablePrefix.value)
    Config.DisableSuffixBox = CBool(chkDisableSuffix.value)
    
    'Debug.Print "Writing: " & IIf(chkTTT.value = 1, "N", "Y")
    Config.AllowMp3Commands = CBool(chkAllowMP3.value)
    Config.MinimizeOnStartup = CBool(chkMinimizeOnStartup.value)
    Config.UseBackupChannel = CBool(chkBackup.value)
    Config.BackupChannel = txtBackupChan.Text
    Config.DoNotUseDirectFList = CBool(chkDoNotUsePacketFList.value)
    Config.DetectUrls = CBool(chkURLDetect.value)
    Config.ShowOfflineFriends = CBool(chkShowOffline.value)
    
    
    For i = 0 To 3
        If optNaming(i).value Then Exit For ' i = index of opt checked
    Next i
    If i = 4 Then i = 0 ' if none were checked, then set to default
    Config.NamespaceConvention = i
    
    Config.UseD2Naming = CBool(chkD2Naming.value)
    Config.ShowStatsIcons = CBool(chkShowUserGameStatsIcons.value)
    Config.ShowFlagIcons = CBool(chkShowUserFlagsIcons.value)
    Config.LogDbAction = CBool(chkLogDBActions.value)
    Config.LogCommands = CBool(chkLogAllCommands.value)
    
    '// Interface Settings
    Config.ChatFont = txtChatFont.Text
    Config.ChatSize = CLng(txtChatSize.Text)
    Config.ChannelFont = txtChanFont.Text
    Config.ChannelSize = CLng(txtChanSize.Text)
    
    '// Idle message settings
    Config.IdlesEnabled = CBool(chkIdles.value)
    Config.IdleDelay = (Val(txtIdleWait.Text) * 2)
    
    Select Case True
        Case optMsg.value:       Config.IdleType = "msg"
        Case optUptime.value:    Config.IdleType = "uptime"
        Case optMP3.value:       Config.IdleType = "mp3"
        Case optQuote.value:     Config.IdleType = "quote"
        Case Else: Config.IdleType = "msg"
    End Select
    
    Config.IdleMessage = txtIdleMsg.Text
    
    '// General Moderation settings
    Config.EnablePhrasebans = CBool(chkPhrasebans.value)
    Config.IpBans = CBool(chkIPBans.value)
    Config.EnforceBanEvasion = CBool(chkBanEvasion.value)
    
    Clients(SC) = "STAR"
    Clients(BW) = "SEXP"
    Clients(D2) = "D2DV"
    Clients(D2X) = "D2XP"
    Clients(W3) = "WAR3"
    Clients(W3X) = "W3XP"
    Clients(W2) = "W2BN"

    For i = 0 To 6
        If (chkCBan(i).value = 1) Then
            If (GetAccess(Clients(i), "GAME").Username = _
                vbNullString) Then
                
                ' redefine array size
                ReDim Preserve DB(UBound(DB) + 1)
                
                With DB(UBound(DB))
                    .Username = Clients(i)
                    .Flags = "B"
                    .ModifiedBy = "(console)"
                    .ModifiedOn = Now
                    .AddedBy = "(console)"
                    .AddedOn = Now
                    .Type = "GAME"
                End With
                
                ' commit modifications
                Call WriteDatabase(GetFilePath(FILE_USERDB))
                
                ' log actions
                If (BotVars.LogDBActions) Then
                    Call LogDbAction(AddEntry, "console", DB(UBound(DB)).Username, "game", _
                        DB(UBound(DB)).Rank, DB(UBound(DB)).Flags)
                End If
            Else
                For j = LBound(DB) To UBound(DB)
                    If ((StrComp(DB(j).Username, Clients(i), vbTextCompare) = 0) And _
                        (StrComp(DB(j).Type, "GAME", vbTextCompare) = 0)) Then
                        
                        If (InStr(1, DB(j).Flags, "B", vbBinaryCompare) = 0) Then
                            With DB(j)
                                .Username = Clients(i)
                                .Flags = "B" & .Flags
                                .ModifiedBy = "(console)"
                                .ModifiedOn = Now
                            End With
                            
                            ' log actions
                            If (BotVars.LogDBActions) Then
                                Call LogDbAction(ModEntry, "console", DB(j).Username, "game", _
                                    DB(j).Rank, DB(j).Flags)
                            End If
                            
                            ' commit modifications
                            Call WriteDatabase(GetFilePath(FILE_USERDB))
                            
                            ' break loop
                            Exit For
                        End If
                    End If
                Next j
            End If
        Else
            If (GetAccess(Clients(i), "GAME").Username <> _
                vbNullString) Then

                For j = LBound(DB) To UBound(DB)
                    If ((StrComp(DB(j).Username, Clients(i), vbTextCompare) = 0) And _
                        (StrComp(DB(j).Type, "GAME", vbTextCompare) = 0)) Then
                        
                        If ((Len(DB(j).Flags) > 1) Or _
                            (DB(j).Rank > 0) Or _
                            (Len(DB(j).Groups) > 1)) Then

                            With DB(j)
                                .Username = Clients(i)
                                .Flags = Replace(.Flags, "B", vbNullString)
                                .ModifiedBy = "(console)"
                                .ModifiedOn = Now
                            End With
                            
                            ' log actions
                            If (BotVars.LogDBActions) Then
                                Call LogDbAction(ModEntry, "console", DB(j).Username, "game", _
                                    DB(j).Rank, DB(j).Flags)
                            End If
                            
                            ' commit modifications
                            Call WriteDatabase(GetFilePath(FILE_USERDB))
                        Else
                            Call RemoveItem(Clients(i), "users", _
                                "GAME")
                                
                            ' log actions
                            If (BotVars.LogDBActions) Then
                                Call LogDbAction(RemEntry, "console", DB(j).Username, "game", _
                                    DB(j).Rank, DB(j).Flags)
                            End If
                            
                            ' reload database entries
                            Call LoadDatabase
                        End If
                        
                        ' break loop
                        Exit For
                    End If
                Next j
            End If
        End If
    Next i
    
    Config.QuietTime = CBool(chkQuiet.value)
    Config.KickOnYell = CBool(chkKOY.value)
    Config.BanUdpPlugs = CBool(chkPlugban.value)
    Config.ChannelProtection = CBool(chkProtect.value)
    Config.ChannelProtectionMessage = txtProtectMsg.Text
    Config.EnforceIdleBans = CBool(chkIdlebans.value)
    Config.KickIdleUsers = CBool(chkIdleKick.value)
    Config.IdleBanDelay = CLng(txtIdleBanDelay.Text)
    Config.BanWc3Peons = CBool(chkPeonbans.value)
    Config.BanUnderLevel = CLng(txtBanW3.Text)
    Config.BanD2UnderLevel = CLng(txtBanD2.Text)
    Config.LevelBanMessage = txtLevelBanMsg.Text

    '// Greet Message Settings
    Config.GreetMessage = txtGreetMsg.Text
    Config.WhisperGreet = CBool(chkWhisperGreet.value)
    Config.UseGreetMessage = CBool(chkGreetMsg.value)
    
    '// Save the config class to disk
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
    
    SaveFontSettings
    
    '// Store cdkeys
    Call WriteCDKeys(cboCDKey)
    
    SaveSettings = True
End Function

' check for potential invalid config stuffs
Public Function InvalidConfigValues() As Boolean

    Dim s As String
    
    If optW3XP.value Or optD2XP.value Then
        If LenB(txtExpKey.Text) = 0 Then
            If optW3XP.value Then
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
        .fileName = vbNullString
        .ShowSave
        If .fileName <> vbNullString Then
            SaveColors .fileName
            MsgBox "ColorList exported.", vbOKOnly
        End If
    End With
End Sub

Private Sub cmdImport_Click()
    With cDLG
        .fileName = vbNullString
        .ShowOpen
        If .fileName <> vbNullString Then
            GetColorLists (.fileName)
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

    Debug.Print "Error " & Err.Number & " (" & Err.description & ") in procedure ShowCurrentColor of Form frmSettings"
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
    txtProxyIP.Enabled = chkUseProxies.value
    txtProxyPort.Enabled = chkUseProxies.value
    optSocks4.Enabled = chkUseProxies.value
    optSocks5.Enabled = chkUseProxies.value
End Sub

Private Sub chkBackup_Click()
    txtBackupChan.Enabled = chkBackup.value
End Sub

Private Sub chkIdlebans_click()
    chkIdleKick.Enabled = chkIdlebans.value
    txtIdleBanDelay.Enabled = chkIdlebans.value
End Sub

Private Sub chkIdles_Click()
    optMsg.Enabled = chkIdles.value
    optUptime.Enabled = chkIdles.value
    optMP3.Enabled = chkIdles.value
    optQuote.Enabled = chkIdles.value
    txtIdleWait.Enabled = chkIdles.value
    txtIdleMsg.Enabled = (optMsg.Enabled And optMsg.value)
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
    chkSpawn.Visible = True
    chkJPN.Visible = True
    cboCDKey.Enabled = True
    txtExpKey.Enabled = False
    chkUseRealm.Enabled = False
    If (chkSHR.value) Then
        SetHashPath GetGamePath("RHSS")
        chkSpawn.Visible = False
        cboCDKey.Enabled = False
    ElseIf (chkJPN.value) Then
        SetHashPath GetGamePath("RTSJ")
    Else
        SetHashPath GetGamePath("RATS")
    End If
    chkUDP.Enabled = True
End Sub

Sub optWAR3_Click()
    chkSHR.Visible = False
    chkSpawn.Visible = False
    chkJPN.Visible = False
    cboCDKey.Enabled = True
    txtExpKey.Enabled = False
    chkUseRealm.Enabled = False
    SetHashPath GetGamePath("3RAW")
    chkUDP.Enabled = False
End Sub

Sub optD2DV_Click()
    chkSHR.Visible = False
    chkSpawn.Visible = False
    chkJPN.Visible = False
    cboCDKey.Enabled = True
    txtExpKey.Enabled = False
    chkUseRealm.Enabled = True
    SetHashPath GetGamePath("VD2D")
    chkUDP.Enabled = False
End Sub

Sub optW2BN_Click()
    chkSHR.Visible = False
    chkSpawn.Visible = True
    chkJPN.Visible = False
    cboCDKey.Enabled = True
    txtExpKey.Enabled = False
    chkUseRealm.Enabled = False
    SetHashPath GetGamePath("NB2W")
    chkUDP.Enabled = True
End Sub

Sub optSEXP_Click()
    chkSHR.Visible = False
    chkSpawn.Visible = False
    chkJPN.Visible = False
    cboCDKey.Enabled = True
    txtExpKey.Enabled = False
    chkUseRealm.Enabled = False
    SetHashPath GetGamePath("RATS")
    chkUDP.Enabled = True
End Sub

Sub optD2XP_Click()
    chkSHR.Visible = False
    chkSpawn.Visible = False
    chkJPN.Visible = False
    cboCDKey.Enabled = True
    txtExpKey.Enabled = True
    chkUseRealm.Enabled = True
    SetHashPath GetGamePath("PX2D")
    chkUDP.Enabled = False
End Sub

Sub optW3XP_Click()
    chkSHR.Visible = False
    chkSpawn.Visible = False
    chkJPN.Visible = False
    cboCDKey.Enabled = True
    txtExpKey.Enabled = True
    chkUseRealm.Enabled = False
    SetHashPath GetGamePath("PX3W")
    chkUDP.Enabled = False
End Sub

Sub optDRTL_Click()
    chkSHR.Visible = True
    chkSpawn.Visible = False
    chkJPN.Visible = False
    cboCDKey.Enabled = False
    txtExpKey.Enabled = False
    chkUseRealm.Enabled = False
    If (chkSHR.value) Then
        SetHashPath GetGamePath("RHSD")
    Else
        SetHashPath GetGamePath("LTRD")
    End If
    chkUDP.Enabled = True
End Sub

Private Sub chkJPN_Click()
    Dim Checked As Boolean
    Checked = CBool(chkJPN.value)
    If (Checked) Then chkSHR.value = vbUnchecked
    If (optSTAR.value) Then
        If (Checked) Then
            SetHashPath GetGamePath("RTSJ")
        Else
            SetHashPath GetGamePath("RATS")
        End If
    End If
End Sub

Private Sub chkSHR_Click()
    Dim Checked As Boolean
    Checked = CBool(chkSHR.value)
    If (Checked) Then chkJPN.value = vbUnchecked
    If (optSTAR.value) Then
        chkSpawn.Visible = Not Checked
        cboCDKey.Enabled = Not Checked
        If (Checked) Then
            SetHashPath GetGamePath("RHSS")
        Else
            SetHashPath GetGamePath("RATS")
        End If
    ElseIf (optDRTL.value) Then
        If (Checked) Then
            SetHashPath GetGamePath("RHSD")
        Else
            SetHashPath GetGamePath("LTRD")
        End If
    End If
End Sub

Private Sub chkGreetMsg_Click()
    chkWhisperGreet.Enabled = chkGreetMsg.value
    txtGreetMsg.Enabled = chkGreetMsg.value
End Sub

Private Sub chkProtect_Click()
    txtProtectMsg.Enabled = chkProtect.value
End Sub

'##########################################
' INIT SUBS
'##########################################

Private Sub InitGenMisc()
    Dim s As String
    Dim i As Integer
    
    chkPAmp.value = Abs(Config.UseProfileAmp)
    chkWhisperCmds.value = Abs(Config.WhisperResponses)
    chkMail.value = Abs(Config.EnableMail)
    
    chkDisablePrefix.value = Abs(Config.DisablePrefixBox)
    chkDisableSuffix.value = Abs(Config.DisableSuffixBox)
    
    'Debug.Print "Loaded value: " & YesToTrue(readcfg(OT, "TTT"), 1) & " converted to " & IIf(YesToTrue(readcfg(OT, "TTT"), 1) = 1, 0, 1)
    
    chkAllowMP3.value = Abs(Config.AllowMp3Commands)
    chkMinimizeOnStartup.value = Abs(Config.MinimizeOnStartup)
    chkShowOffline.value = Abs(Config.ShowOfflineFriends)
    
    optNaming(Config.NamespaceConvention).value = True
    chkD2Naming.value = Abs(Config.UseD2Naming)
    
    chkShowUserGameStatsIcons.value = Abs(Config.ShowStatsIcons)
    chkShowUserFlagsIcons.value = Abs(Config.ShowFlagIcons)
    
    chkURLDetect.value = Abs(Config.DetectUrls)
    chkDoNotUsePacketFList.value = Abs(Config.DoNotUseDirectFList)
          
    chkBackup.value = Abs(Config.UseBackupChannel)
    Call chkBackup_Click
    
    txtBackupChan.Text = Config.BackupChannel
    
    chkLogDBActions.value = Abs(Config.LogDbAction)
    chkLogAllCommands.value = Abs(Config.LogCommands)
End Sub

Private Sub InitGenIdles()
    Dim s As String
    
    txtIdleWait.Text = Config.IdleDelay / 2
    
    Select Case Config.IdleType
        Case "msg", vbNullString
            optMsg.value = True
            Call optMsg_Click
        Case "quote"
            optQuote.value = True
            Call optQuote_Click
        Case "uptime"
            optUptime.value = True
            Call optUptime_Click
        Case "mp3"
            optMP3.value = True
            Call optMP3_Click
        Case Else
            optMsg.value = True
            Call optMsg_Click
    End Select
    
    txtIdleMsg.Text = Config.IdleMessage
    If LenB(txtIdleMsg.Text) = 0 Then txtIdleMsg.Text = "/me is a %v by Stealth - http://www.stealthbot.net"
    
    chkIdles.value = Abs(Config.IdlesEnabled)
    Call chkIdles_Click
    
End Sub

Private Sub InitGenGreets()
    txtGreetMsg.Text = Config.GreetMessage
    chkGreetMsg.value = Abs(Config.UseGreetMessage)
    Call chkGreetMsg_Click
    
    chkWhisperGreet.value = Abs(Config.WhisperGreet)
End Sub

Private Sub InitBasicConfig()
    Dim s As String
    Dim f As Integer
    Dim i As Integer
    Dim AddCurrent As Boolean
    Dim Item As String
    Dim AdditionalServerList As Collection
    
    txtUsername.Text = Config.Username
    txtPassword.Text = Config.Password
    cboCDKey.Text = Config.CdKey
    
    txtExpKey.Text = Config.ExpKey
    
    txtHomeChan.Text = Config.HomeChannel
    txtOwner.Text = Config.BotOwner
    
    OldBotOwner = txtOwner.Text

    With cboServer
        .Text = Config.Server
        
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
        
        For i = 0 To .ListCount
            Item = .List(i)
            
            If StrComp(Item, s, vbBinaryCompare) = 0 Then
            
                AddCurrent = False
                
            End If
        Next i
        
        ' if not, add it (first)
        If AddCurrent Then
        
            .AddItem s, 0
            
        End If
    End With
    
    txtTrigger.Text = Config.Trigger
    
    s = Config.Product
    Select Case StrReverse(UCase(s))
        Case "STAR":    Call optSTAR_Click: optSTAR.value = True: chkSHR.value = vbUnchecked: chkJPN.value = vbUnchecked
        Case "SEXP":    Call optSEXP_Click: optSEXP.value = True
        Case "D2DV":    Call optD2DV_Click: optD2DV.value = True
        Case "D2XP":    Call optD2XP_Click: optD2XP.value = True
        Case "W2BN":    Call optW2BN_Click: optW2BN.value = True
        Case "WAR3":    Call optWAR3_Click: optWAR3.value = True
        Case "W3XP":    Call optW3XP_Click: optW3XP.value = True
        Case "DRTL":    Call optDRTL_Click: optDRTL.value = True: chkSHR.value = vbUnchecked
        Case "DSHR":    Call optDRTL_Click: optDRTL.value = True: chkSHR.value = vbChecked
        Case "SSHR":    Call optSTAR_Click: optSTAR.value = True: chkSHR.value = vbChecked ' unchecks jpn
        Case "JSTR":    Call optSTAR_Click: optSTAR.value = True: chkJPN.value = vbChecked ' unchecks shr
        Case Else:      Call optSTAR_Click: optSTAR.value = True: chkSHR.value = vbUnchecked: chkJPN.value = vbUnchecked
    End Select
    
    chkSpawn.value = Abs(Config.UseSpawnKey)
    
    chkUseRealm.value = Abs(Config.UseRealm)
    
    Call LoadCDKeys(cboCDKey)
    
End Sub

Private Sub InitGenMod()
    'Dim s As String
    
    chkPhrasebans.value = Abs(Config.EnablePhrasebans)
    chkIPBans.value = Abs(Config.IpBans)
    chkQuiet.value = Abs(Config.QuietTime)
    chkKOY.value = Abs(Config.KickOnYell)
    chkPlugban.value = Abs(Config.BanUdpPlugs)
    chkPeonbans.value = Abs(Config.BanWc3Peons)

    chkBanEvasion.value = Abs(Config.EnforceBanEvasion)
    
    chkProtect.value = Abs(Config.ChannelProtection)
    Call chkProtect_Click
    
    txtProtectMsg.Text = Config.ChannelProtectionMessage
    
    chkIdlebans.value = Abs(Config.EnforceIdleBans)
    chkIdleKick.value = Abs(Config.KickIdleUsers)
    Call chkIdlebans_click
    
    txtIdleBanDelay.Text = Config.IdleBanDelay
    
    ' grab client ban settings from database
    
    If (InStr(1, GetAccess("STAR", "GAME").Flags, "B", _
        vbBinaryCompare) <> 0) Then
    
        chkCBan(SC).value = 1
    End If
    
    If (InStr(1, GetAccess("SEXP", "GAME").Flags, "B", _
        vbBinaryCompare) <> 0) Then
    
        chkCBan(BW).value = 1
    End If
    
    If (InStr(1, GetAccess("D2DV", "GAME").Flags, "B", _
        vbBinaryCompare) <> 0) Then
    
        chkCBan(D2).value = 1
    End If
    
    If (InStr(1, GetAccess("D2XP", "GAME").Flags, "B", _
        vbBinaryCompare) <> 0) Then
    
        chkCBan(D2X).value = 1
    End If
    
    If (InStr(1, GetAccess("W2BN", "GAME").Flags, "B", _
        vbBinaryCompare) <> 0) Then
    
        chkCBan(W2).value = 1
    End If
    
    If (InStr(1, GetAccess("WAR3", "GAME").Flags, "B", _
        vbBinaryCompare) <> 0) Then
    
        chkCBan(W3).value = 1
    End If
    
    If (InStr(1, GetAccess("W3XP", "GAME").Flags, "B", _
        vbBinaryCompare) <> 0) Then
    
        chkCBan(W3X).value = 1
    End If

    txtLevelBanMsg.Text = Config.LevelBanMessage
    If LenB(txtLevelBanMsg.Text) = 0 Then txtLevelBanMsg.Text = "You are below the required level for entry."
    
    txtBanD2.Text = Config.BanD2UnderLevel
    txtBanW3.Text = Config.BanUnderLevel
End Sub

Private Sub InitGenInterface()
    Dim s As String
    
    chkJoinLeaves.value = Abs(Config.ShowJoinLeaves)
    chkFilter.value = Abs(Config.UseChatFilters)
    chkSplash.value = Abs(Config.ShowSplashScreen)
    chkUTF8.value = Abs(Config.UTF8)
    chkFlash.value = Abs(Config.FlashOnEvents)
    chkNoTray.value = Abs(Not Config.MinimizeToTray)
    chkNoAutocomplete.value = Abs(Config.DisableAutoComplete)
    chkNoColoring.value = Abs(Config.DisableNameColoring)
    
    Select Case Config.LoggingLevel
        Case 0:
            cboLogging.ListIndex = 2
        Case 1:
            cboLogging.ListIndex = 1
        Case Else:
            cboLogging.ListIndex = 0
    End Select
    
    cboTimestamp.ListIndex = Config.Timestamp
    
    txtMaxBackLogSize.Text = Config.MaxBacklogSize
    
    txtMaxLogSize.Text = Config.MaxLogFileSize
    
End Sub

Private Sub InitFontsColors()
    'Dim s As String
    
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

Private Sub InitConnAdvanced()
    Dim s As String
    
    ' Connection method
    cboConnMethod.ListIndex = CInt(Abs(Not Config.UseBnls))
    
    ' Set selected BNLS server
    If Config.UseBnlsFinder Or Len(Config.BnlsServer) = 0 Then
        cboBNLSServer.ListIndex = 0
    ElseIf Len(Config.BnlsServer) > 0 Then
        cboBNLSServer.ListIndex = GetBnlsIndex(Config.BnlsServer)
    End If
    
    txtEmail.Text = Config.RegisterEmailDefault
    
    cboSpoof.ListIndex = Config.PingSpoofing
    chkUDP.value = Abs(Config.UseUDP)
    
    chkConnectOnStartup.value = Abs(Config.ConnectOnStartup)
    txtReconDelay.Text = Config.ReconnectDelay
    If LenB(txtReconDelay.Text) = 0 Then txtReconDelay.Text = 1000
    
    chkUseProxies.value = Abs(Config.UseProxy)
    Call chkUseProxies_Click
    
    txtProxyPort.Text = Config.ProxyPort
    txtProxyIP.Text = Config.ProxyIP
    
    If Config.ProxyIsSocks5 Then
        optSocks5.value = True
        optSocks4.value = False
    Else
        optSocks5.value = False
        optSocks4.value = True
    End If
    
    ' Adjust "BNLS server" label 2 pixels down
    lbl5(12).Top = lbl5(12).Top + (2 * Screen.TwipsPerPixelY)
End Sub

Private Sub InitAllPanels()
    InitGenMisc
    InitConnAdvanced
    InitFontsColors
    InitGenInterface
    InitGenMod
    InitBasicConfig
    InitGenGreets
    InitGenIdles
    
    If Config.FileExists Then
        ShowPanel spConnectionConfig
        FirstRun = 1
    End If
End Sub

Private Function YesToTrue(ByVal s As String, ByVal bDefault As Integer) As Integer
    YesToTrue = 0
    
    If LenB(s) > 0 Then
        If StrComp(UCase(s), "Y", vbBinaryCompare) = 0 Then
            YesToTrue = 1
        End If
    Else
        YesToTrue = bDefault
    End If
End Function

Private Function Cv(ByVal i As Integer) As String
    Select Case i
        Case 0: Cv = "N"
        Case 1: Cv = "Y"
    End Select
End Function

Function CDKeyReplacements(ByVal inString As String) As String
    inString = Replace(inString, "-", "")
    inString = Replace(inString, " ", "")
    CDKeyReplacements = Trim$(inString)
End Function

Sub SaveFontSettings()
    Dim ResizeChatElements As Boolean
    Dim ResizeChannelElements As Boolean
    
    If (StrComp(InitChatFont, txtChatFont.Text, vbTextCompare)) Then
        Config.ChatFont = txtChatFont.Text
        frmChat.rtbChat.Font.Name = txtChatFont.Text
        frmChat.cboSend.Font.Name = txtChatFont.Text
        frmChat.txtPre.Font.Name = txtChatFont.Text
        frmChat.txtPost.Font.Name = txtChatFont.Text
        frmChat.rtbWhispers.Font.Name = txtChatFont.Text
        ResizeChatElements = True
    End If
    
    If Not InitChatSize = CInt(txtChatSize.Text) Then
        Config.ChatSize = txtChatSize.Text
        frmChat.rtbChat.Font.Size = CInt(txtChatSize.Text)
        frmChat.cboSend.Font.Size = CInt(txtChatSize.Text)
        frmChat.txtPre.Font.Size = CInt(txtChatSize.Text)
        frmChat.txtPost.Font.Size = CInt(txtChatSize.Text)
        frmChat.rtbWhispers.Font.Size = CInt(txtChatSize.Text)
        ResizeChatElements = True
    End If
    
    If (StrComp(InitChanFont, txtChanFont.Text, vbTextCompare)) Then
        Config.ChannelFont = txtChanFont.Text
        frmChat.lvChannel.Font.Name = txtChanFont.Text
        frmChat.lvClanList.Font.Name = txtChanFont.Text
        frmChat.lvFriendList.Font.Name = txtChanFont.Text
        frmChat.ListviewTabs.Font.Name = txtChanFont.Text
        frmChat.lblCurrentChannel.Font.Name = txtChanFont.Text
        ResizeChannelElements = True
    End If
    
    If Not InitChanSize = CInt(txtChanSize.Text) Then
        Config.ChannelSize = txtChanSize.Text
        frmChat.lvChannel.Font.Size = CInt(txtChanSize.Text)
        frmChat.lvClanList.Font.Size = CInt(txtChanSize.Text)
        frmChat.lvFriendList.Font.Size = CInt(txtChanSize.Text)
        frmChat.ListviewTabs.Font.Size = CInt(txtChanSize.Text)
        frmChat.lblCurrentChannel.Font.Size = CInt(txtChanSize.Text)
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
        frmChat.Form_Resize
    End If
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



