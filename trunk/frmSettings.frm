VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{CA5A8E1E-C861-4345-8FF8-EF0A27CD4236}#1.0#0"; "vbalTreeView6.ocx"
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
      TabIndex        =   179
      Text            =   "Profile Selector"
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton cmdWebsite 
      Caption         =   "&Website"
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
      TabIndex        =   106
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton cmdReadme 
      Caption         =   "&Readme"
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
      TabIndex        =   105
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
      TabIndex        =   107
      Top             =   4920
      Width           =   2535
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Apply and Cl&ose"
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
      TabIndex        =   108
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
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
      TabIndex        =   109
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
      Index           =   2
      Left            =   3000
      TabIndex        =   98
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
         TabIndex        =   187
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
         TabIndex        =   181
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
         TabIndex        =   180
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
         TabIndex        =   26
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
         TabIndex        =   21
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
         TabIndex        =   25
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
         TabIndex        =   24
         Text            =   "(Select One)"
         Top             =   960
         Width           =   3135
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
         TabIndex        =   23
         Text            =   "0"
         Top             =   4410
         Width           =   735
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   22
         Text            =   "(Select One)"
         Top             =   3000
         Width           =   5895
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
         TabIndex        =   189
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
         TabIndex        =   188
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
         TabIndex        =   141
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
         TabIndex        =   133
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
         TabIndex        =   132
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
         TabIndex        =   131
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
         TabIndex        =   130
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
         TabIndex        =   124
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
      TabIndex        =   120
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
         TabIndex        =   177
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
         TabIndex        =   32
         Text            =   "cboColorList"
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
         TabIndex        =   33
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
         TabIndex        =   34
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
         TabIndex        =   88
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
         TabIndex        =   89
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
         TabIndex        =   90
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
         TabIndex        =   93
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
         TabIndex        =   94
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
         TabIndex        =   31
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
         TabIndex        =   95
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
         TabIndex        =   91
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
         TabIndex        =   92
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
         TabIndex        =   29
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
         TabIndex        =   27
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
         TabIndex        =   30
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
         TabIndex        =   28
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
         TabIndex        =   150
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label lblEg 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3360
         TabIndex        =   149
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
         TabIndex        =   148
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
         TabIndex        =   147
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
         TabIndex        =   146
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
         TabIndex        =   145
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
         TabIndex        =   144
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
         TabIndex        =   143
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
         TabIndex        =   142
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
         TabIndex        =   140
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
         TabIndex        =   139
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
         TabIndex        =   138
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
         TabIndex        =   137
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
         TabIndex        =   136
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
         TabIndex        =   135
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
         TabIndex        =   134
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
         TabIndex        =   178
         Top             =   4200
         Width           =   5895
      End
   End
   Begin VB.Frame fraPanel 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   4815
      Index           =   8
      Left            =   3000
      TabIndex        =   122
      Top             =   0
      Width           =   6615
      Begin VB.CheckBox chkShowUserGameStatsIcons 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Show user game stats icons"
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
         Height          =   375
         Left            =   360
         TabIndex        =   193
         Top             =   3680
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CheckBox chkGameConventions 
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
         Left            =   3840
         TabIndex        =   192
         Top             =   4260
         Width           =   2175
      End
      Begin VB.CheckBox chkGameConventions 
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
         Index           =   1
         Left            =   3840
         TabIndex        =   191
         Top             =   4005
         Width           =   2175
      End
      Begin VB.CheckBox chkGameConventions 
         BackColor       =   &H00000000&
         Caption         =   "Use Game-Specific Username Conventions"
         BeginProperty Font 
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
         Index           =   0
         Left            =   3480
         TabIndex        =   190
         ToolTipText     =   "Disable the use of legacy naming conventions for chat interface and command usage"
         Top             =   3560
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
         TabIndex        =   39
         ToolTipText     =   "Disables the smaller suffix box to the right of the box you type in to send text to Battle.net"
         Top             =   2280
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
         TabIndex        =   38
         ToolTipText     =   "Disables the smaller prefix box to the left of the box you type in to send text to Battle.net"
         Top             =   1920
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
         TabIndex        =   184
         ToolTipText     =   "Any commands issued to the bot will be logged in a file in the bot's Logs folder called 'commandlog.txt'."
         Top             =   2250
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
         TabIndex        =   183
         ToolTipText     =   "Any actions that change the database will be logged in the log folder in a file called 'database.txt'."
         Top             =   1920
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
         TabIndex        =   41
         ToolTipText     =   "Determines whether or not offline friends are hidden from /f list"
         Top             =   3000
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
         TabIndex        =   40
         ToolTipText     =   "Enables automatic URL detection and highlighting in the main chat window."
         Top             =   2640
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
         TabIndex        =   42
         ToolTipText     =   "Disable the internal friends' list (alternative channel list)"
         Top             =   3310
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
         MaxLength       =   25
         TabIndex        =   87
         Top             =   3140
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
         TabIndex        =   45
         ToolTipText     =   "Automatically minimize on startup."
         Top             =   1560
         Width           =   2535
      End
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
         Left            =   3480
         TabIndex        =   44
         ToolTipText     =   "Automatically connect when the bot starts up."
         Top             =   1200
         Width           =   2535
      End
      Begin VB.CheckBox chkAllowMP3 
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
         Left            =   3480
         TabIndex        =   43
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
         TabIndex        =   36
         ToolTipText     =   "Whisper the return messages of all bot commands."
         Top             =   1200
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
         TabIndex        =   35
         ToolTipText     =   "Enable/disable ProfileAmp - writes Winamp's currently played song to your profile every 30 seconds"
         Top             =   840
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
         TabIndex        =   37
         ToolTipText     =   "Enable/disable checking of the mail.ini file when people join."
         Top             =   1560
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
         TabIndex        =   86
         ToolTipText     =   "The bot will join a specified channel when kicked, instead of rejoining."
         Top             =   2520
         Width           =   2895
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
         TabIndex        =   171
         Top             =   2910
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
         TabIndex        =   167
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
      TabIndex        =   97
      Top             =   0
      Width           =   6615
      Begin VB.CheckBox chkBNLSAlt 
         BackColor       =   &H00000000&
         Caption         =   "Enable Automatic BNLS Server Finder"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   194
         Top             =   4440
         Width           =   4695
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
         Left            =   2760
         List            =   "frmSettings.frx":00D8
         TabIndex        =   47
         Text            =   "cboBNLSServer"
         Top             =   1200
         Width           =   3495
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
         TabIndex        =   50
         Text            =   "1000"
         Top             =   3360
         Width           =   855
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
         TabIndex        =   55
         Top             =   3360
         Width           =   855
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
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   3360
         Width           =   855
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
         Left            =   4440
         MaxLength       =   5
         TabIndex        =   53
         Top             =   3000
         Width           =   1815
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
         Left            =   4440
         MaxLength       =   15
         TabIndex        =   52
         Top             =   2640
         Width           =   1815
      End
      Begin VB.CheckBox chkUseProxies 
         BackColor       =   &H00000000&
         Caption         =   "Use a proxy to connect"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   51
         ToolTipText     =   "Routes your Battle.net and/or BNLS connection through a SOCKS4 or 5 proxy."
         Top             =   2280
         Width           =   2055
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
         Left            =   360
         TabIndex        =   49
         ToolTipText     =   "Sets whether or not you have a lag plug when you sign on. If you don't know what this is, leave it off."
         Top             =   2880
         Width           =   1455
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
         TabIndex        =   48
         Text            =   "cboSpoof"
         Top             =   2520
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
         TabIndex        =   46
         Top             =   840
         Width           =   4335
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
         TabIndex        =   186
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label lbl5 
         BackColor       =   &H00000000&
         Caption         =   "(in milliseconds)"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         Left            =   480
         TabIndex        =   175
         Top             =   3480
         Width           =   1215
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
         TabIndex        =   174
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label lbl5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Proxy port"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   170
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label lbl5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Proxy IP address"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   169
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         X1              =   2760
         X2              =   2760
         Y1              =   2280
         Y2              =   3720
      End
      Begin VB.Label lbl5 
         BackColor       =   &H00000000&
         Caption         =   $"frmSettings.frx":00DA
         BeginProperty Font 
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
         Index           =   1
         Left            =   360
         TabIndex        =   129
         Top             =   3840
         Width           =   5895
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
         Left            =   360
         TabIndex        =   128
         Top             =   2280
         Width           =   1575
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
         Height          =   255
         Left            =   360
         TabIndex        =   127
         Top             =   1920
         Width           =   6015
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
         TabIndex        =   126
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label lbl5 
         BackColor       =   &H00000000&
         Caption         =   "Local hashing supports all game clients. Your current hash file path is:"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   125
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
         TabIndex        =   123
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
      TabIndex        =   101
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
         TabIndex        =   58
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
         TabIndex        =   61
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
         TabIndex        =   59
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
         TabIndex        =   60
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
         TabIndex        =   56
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
         TabIndex        =   62
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
         TabIndex        =   57
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
         TabIndex        =   166
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
         TabIndex        =   165
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
         TabIndex        =   164
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
         TabIndex        =   162
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
         TabIndex        =   163
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
      TabIndex        =   100
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
         TabIndex        =   64
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
         TabIndex        =   65
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
         TabIndex        =   63
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
         Height          =   1935
         Left            =   360
         TabIndex        =   161
         Top             =   2280
         Width           =   5895
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
         TabIndex        =   159
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
         TabIndex        =   158
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
      TabIndex        =   121
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
         TabIndex        =   182
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
         TabIndex        =   176
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
         TabIndex        =   75
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
         TabIndex        =   85
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
         TabIndex        =   82
         Top             =   2520
         Width           =   1695
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
         TabIndex        =   81
         Top             =   2280
         Width           =   1695
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
         TabIndex        =   66
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
         TabIndex        =   67
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
         TabIndex        =   76
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
         TabIndex        =   80
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
         TabIndex        =   77
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
         TabIndex        =   79
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
         TabIndex        =   78
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
         TabIndex        =   68
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
         TabIndex        =   72
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
         TabIndex        =   71
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
         TabIndex        =   69
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
         TabIndex        =   84
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
         TabIndex        =   70
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
         TabIndex        =   83
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
         TabIndex        =   73
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
         TabIndex        =   74
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
         TabIndex        =   168
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
         TabIndex        =   157
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
         TabIndex        =   156
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
         TabIndex        =   155
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
         TabIndex        =   154
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
         TabIndex        =   153
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
         TabIndex        =   152
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
         TabIndex        =   151
         Top             =   240
         Width           =   4815
      End
   End
   Begin VB.Frame fraPanel 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   4815
      Index           =   4
      Left            =   3000
      TabIndex        =   99
      Top             =   0
      Width           =   6615
   End
   Begin VB.Frame fraPanel 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   4815
      Index           =   9
      Left            =   3000
      TabIndex        =   102
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
         TabIndex        =   104
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
         TabIndex        =   103
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
      TabIndex        =   96
      Top             =   0
      Width           =   6615
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
         Left            =   840
         TabIndex        =   6
         ToolTipText     =   "This account has full control over the bot. Use carefully!"
         Top             =   3960
         Width           =   1935
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
      Begin VB.OptionButton optCHAT 
         BackColor       =   &H00000000&
         Caption         =   "Telnet Client"
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
         Height          =   255
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Not yet supported."
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
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   3360
         Width           =   1455
      End
      Begin VB.OptionButton optWAR3 
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
         Height          =   375
         Index           =   19
         Left            =   240
         TabIndex        =   185
         Top             =   3840
         Width           =   615
      End
      Begin VB.Label lblAddCurrentKey 
         BackColor       =   &H00000000&
         Caption         =   "(add current)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   960
         TabIndex        =   173
         Top             =   2040
         Width           =   975
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
         TabIndex        =   172
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
         TabIndex        =   119
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
         TabIndex        =   118
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
         TabIndex        =   116
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
         TabIndex        =   115
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
         TabIndex        =   114
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
         TabIndex        =   113
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
         TabIndex        =   112
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
         TabIndex        =   111
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
         TabIndex        =   110
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
         TabIndex        =   117
         Top             =   240
         Width           =   4815
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   4320
      TabIndex        =   160
      Top             =   2160
      Width           =   1215
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
Private Const MN            As String = "Main"
Private Const OT            As String = "Other"
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
Const W3X    As Byte = 5
Const W2   As Byte = 6

Const BNLS_SERVER_COUNT = 2

Private Enum enuConfigSection
    secMain = 0
    secOther = 1
End Enum

Private Sub chkGameConventions_Click(Index As Integer)
    If (Index = 0) Then
        If (chkGameConventions(Index).Value = 0) Then
            chkGameConventions(1).Enabled = False
            chkGameConventions(2).Enabled = False
        Else
            chkGameConventions(1).Enabled = True
            chkGameConventions(2).Enabled = True
        End If
    End If
End Sub

Private Sub Form_Load()
    Me.Icon = frmChat.Icon
    
    Dim nRoot As cTreeViewNode
    Dim nCurrent As cTreeViewNode
    Dim nTopLevel As cTreeViewNodes
    Dim nOptLevel As cTreeViewNodes
    Dim lMouseOver As Long
    Dim s As String
    Dim serverList() As String
    Dim I As Long, j As Long, k As Long
    
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
                        nOptLevel.Add , etvwChild, "conn_config", "General Settings"  'general setup
                        nOptLevel.Add , etvwChild, "conn_advanced", "Advanced Settings"     'proxies/spoofing/bnls
                        
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
                        nOptLevel.Add , etvwChild, "op_moderation", "Moderation Settings"
                        nOptLevel.Add , etvwChild, "op_greets", "Greet Message Settings"
                        nOptLevel.Add , etvwChild, "op_idles", "Idle Message Settings"
                        nOptLevel.Add , etvwChild, "op_misc", "Miscellaneous Settings"
                        
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
        .AddItem "No ping spoofing"
        .AddItem "Spoof 0ms ping"
        .AddItem "Spoof -1ms ping"
    End With
    
    With cboConnMethod
        .AddItem "BNLS - Battle.net Logon Server"
        .AddItem "ADVANCED - Local hashing"
        .ListIndex = 0
    End With
    
    With cboBNLSServer
        .AddItem "jbls.org"
        .AddItem "bnls.valhallalegends.com"
        
        If BotVars.BNLSServer = "jbls.org" Then
            .ListIndex = 0
        ElseIf Len(BotVars.BNLSServer) > 0 And BotVars.BNLSServer <> "bnls.valhallalegends.com" Then
            .AddItem BotVars.BNLSServer
            .ListIndex = BNLS_SERVER_COUNT
        Else
            .ListIndex = 0
        End If
        
        If LenB(Dir$(GetFilePath("AdditionalBNLSservers.txt"))) > 0 Then
            With cboBNLSServer
                I = FreeFile
                
                Open GetFilePath("AdditionalBNLSservers.txt") For Input As #I
                    While Not EOF(I)
                        Line Input #I, s
                        
                        For j = 0 To .ListCount
                            If StrComp(.List(j), s, vbTextCompare) = 0 Then
                                j = -1
                                Exit For
                            End If
                        Next j
                        
                        If j >= 0 Then
                            .AddItem s
                        End If
                    Wend
                Close #I
            End With
        End If
        
        ' Get some servers from the status page. (these should always be online)
        s = vbNullString
        If (LenB(ReadCfg("Override", "BNLSSource")) > 0) Then
            s = frmChat.INet.OpenURL(ReadCfg("Override", "BNLSSource"))
        End If
        If ((LenB(s) > 0) And (Right(s, 2) = vbCrLf)) Then
            serverList = Split(s, vbCrLf)
            For k = 0 To (UBound(serverList) - 1)
                s = serverList(k)
                
                ' Why doesn't combo box have an exists() method? ><
                For j = 0 To .ListCount
                    If StrComp(.List(j), s, vbTextCompare) = 0 Then
                        j = -1
                        Exit For
                    End If
                Next j
                        
                If j >= 0 Then
                    .AddItem s
                End If
            Next
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
    
    s = ReadCfg("Position", "LastSettingsPanel")
    
    If LenB(s) > 0 Then
        If StrictIsNumeric(s) Then
            'Debug.Print "Reading: " & s
            
            If Val(s) < 10 Then
                ShowPanel CInt(s)
            Else
                ShowPanel spSplash, 1
            End If
        Else
            ShowPanel spSplash, 1
        End If
    Else
        ShowPanel spSplash, 1
    End If
End Sub

Private Sub Form_GotFocus()
    If Len(cboCDKey.Text) = 0 And cboCDKey.ListCount > 0 Then
        cboCDKey.ListIndex = 1
    End If
End Sub

Function KeyToIndex(ByVal sKey As String) As Byte
    
    Select Case sKey
        Case "splash":          KeyToIndex = 9
    
        Case "conn_config":     KeyToIndex = 0
        Case "conn_advanced":   KeyToIndex = 1
        
        Case "int_general":     KeyToIndex = 2
        Case "int_fonts":       KeyToIndex = 3
    
        Case "op_moderation":   KeyToIndex = 5
        Case "op_greets":       KeyToIndex = 6
        Case "op_idles":        KeyToIndex = 7
        Case "op_misc":         KeyToIndex = 8
        
        Case Else:              KeyToIndex = 9
    End Select
    
End Function

Sub ShowPanel(ByVal Index As enuSettingsPanels, Optional ByVal Mode As Byte = 0)

    Static ActivePanel As Integer
    
    If PanelsInitialized Then
        If Mode = 1 Then
            fraPanel(KeyToIndex("splash")).ZOrder vbBringToFront
            ActivePanel = KeyToIndex("splash")
        Else
            'fraPanel(ActivePanel).ZOrder vbSendToBack
            fraPanel(Index).ZOrder vbBringToFront
            ActivePanel = Index
            WriteINI "Position", "LastSettingsPanel", ActivePanel
            
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

Private Sub cboConnMethod_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboLogging_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboSpoof_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboTimestamp_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboColorList_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboBNLSServer_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub cmdStepByStep_Click()
    frmCustomInputBox.Show
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdSave_Click
End Sub

Private Sub cmdWebsite_Click()
    Call frmChat.mnuHelpWebsite_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call frmChat.DeconstructSettings
    Set colProfiles = Nothing
End Sub

Sub lblAddCurrentKey_Click()
    Dim I As Integer
    Dim s As String
    
    s = CDKeyReplacements(cboCDKey.Text)
    
    If cboCDKey.ListCount > -1 Then
        For I = 0 To cboCDKey.ListCount
            If StrComp(cboCDKey.List(I), s, vbTextCompare) = 0 Then
                Exit Sub
            End If
        Next I
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
        ShowPanel KeyToIndex(tvw.SelectedItem.Key)
    Else
        FirstRun = 0
    End If
End Sub

Private Sub WINI(ByVal sKey As String, ByVal sVal As String, ByVal Section As enuConfigSection)
    Select Case Section
        Case secMain
            WriteINI "Main", sKey, sVal
        
        Case secOther
            WriteINI "Other", sKey, sVal
    End Select
End Sub

Private Function DoCDKeyLengthCheck(ByVal sKey As String, ByVal sProd As String) As Boolean
    sKey = CDKeyReplacements(sKey)
    
    DoCDKeyLengthCheck = True
    
    Select Case sProd
        Case "STAR", "SEXP"
            If ((Len(sKey) <> 13) And (Len(sKey) <> 26)) Then DoCDKeyLengthCheck = False
            
        Case "D2DV", "D2XP"
            If ((Len(sKey) <> 16) And (Len(sKey) <> 26)) Then DoCDKeyLengthCheck = False
            
        Case "W2BN"
            If (Len(sKey) <> 16) Then DoCDKeyLengthCheck = False
            
        Case "WAR3", "W3XP"
            If (Len(sKey) <> 26) Then DoCDKeyLengthCheck = False
                
        Case Else
            DoCDKeyLengthCheck = False
    End Select
End Function

Private Function SaveSettings() As Boolean
    Dim s As String
    Dim Clients(6) As String
    Dim I As Long, j As Long
    
    '// First, CDKey Length check and corresponding stuff that needs to run first:
    Select Case True
        Case optSTAR.Value: s = "STAR"
        Case optSEXP.Value: s = "SEXP"
        Case optD2DV.Value: s = "D2DV"
        Case optD2XP.Value: s = "D2XP"
        Case optWAR3.Value: s = "WAR3"
        Case optW3XP.Value: s = "W3XP"
        Case optW2BN.Value: s = "W2BN"
        Case optCHAT.Value: s = "CHAT"
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
    
    WINI "Product", StrReverse(s), secMain
    
    '// The rest of the basic config now
    WINI "Username", txtUsername.Text, secMain
    WINI "Password", txtPassword.Text, secMain
    WINI "CDKey", CDKeyReplacements(cboCDKey.Text), secMain
    WINI "ExpKey", CDKeyReplacements(txtExpKey.Text), secMain
    WINI "LODKey", vbNullString, secMain
    WINI "HomeChan", txtHomeChan.Text, secMain
    WINI "Server", cboServer.Text, secMain
    WINI "BNLSServer", cboBNLSServer.Text, secMain
    
    With cboBNLSServer
        j = -1
        
        For j = 0 To .ListCount
            If StrComp(.Text, .List(j)) = 0 Then
                j = -1
                Exit For
            End If
        Next j
        
        If j >= 0 Or .ListCount > BNLS_SERVER_COUNT Then
            I = FreeFile
            
            Open GetFilePath("AdditionalBNLSservers.txt") For Output As #I
                s = cboBNLSServer.Text
                        
                Print #I, .Text
                
                For j = BNLS_SERVER_COUNT To .ListCount
                    Print #I, .List(j)
                Next j
            Close #I
        End If
    End With
    ' there is WINI and WriteINI? =\
    WINI "Trigger", "{" & txtTrigger.Text & "}", secMain
    WINI "ConfigVersion", CONFIG_VERSION, secMain
    WINI "Owner", txtOwner.Text, secMain
    
    WINI "UseRealm", Cv(chkUseRealm.Value), secMain
    
    '// General Interface settings
    WINI "JoinLeaves", Cv(chkJoinLeaves.Value), secOther
    WINI "Filters", Cv(chkFilter.Value), secOther
    WINI "ShowSplash", Cv(chkSplash.Value), secMain
    WINI "FlashWindow", Cv(chkFlash.Value), secOther
    WINI "Timestamp", cboTimestamp.ListIndex, secOther
    WINI "UTF8", Cv(chkUTF8.Value), secMain
    
    If (cboLogging.ListIndex = 0) Then
        WINI "Logging", 2, secOther
    ElseIf (cboLogging.ListIndex = 1) Then
        WINI "Logging", 1, secOther
    ElseIf (cboLogging.ListIndex = 2) Then
        WINI "Logging", 0, secOther
    End If
    
    WINI "MaxBacklogSize", txtMaxBackLogSize.Text, secMain
    WINI "MaxLogFileSize", txtMaxLogSize.Text, secMain
    WINI "NoTray", Cv(chkNoTray.Value), secOther
    WINI "NoAutocomplete", Cv(chkNoAutocomplete.Value), secOther
    WINI "NoColoring", Cv(chkNoColoring.Value), secOther
    
    '// Misc General Settings
    WINI "ProfileAmp", Cv(chkPAmp.Value), secOther
    WINI "WhisperBack", Cv(chkWhisperCmds.Value), secMain
    WINI "Mail", Cv(chkMail.Value), secOther
    WINI "DisablePrefix", Cv(chkDisablePrefix.Value), secOther
    WINI "DisableSuffix", Cv(chkDisableSuffix.Value), secOther
    
    'Debug.Print "Writing: " & IIf(chkTTT.value = 1, "N", "Y")
'    WINI "DisableMonitor", Cv(chkDisableMonitor.Value), secOther
    WINI "AllowMP3", Cv(chkAllowMP3.Value), secOther
    WINI "ConnectOnStartup", Cv(chkConnectOnStartup.Value), secMain
    WINI "MinimizeOnStartup", Cv(chkMinimizeOnStartup.Value), secMain
    WINI "UseBackupChan", Cv(chkBackup.Value), secMain
    WINI "BackupChan", txtBackupChan.Text, secMain
    WINI "DoNotUseDirectFList", Cv(chkDoNotUsePacketFList.Value), secMain
    WINI "URLDetect", Cv(chkURLDetect.Value), secMain
    WINI "ShowOfflineFriends", Cv(chkShowOffline.Value), secMain
    WINI "UseGameConventions", Cv(chkGameConventions(0).Value), secOther
    WINI "UseD2GameConventions", Cv(chkGameConventions(1).Value), secOther
    WINI "UseW3GameConventions", Cv(chkGameConventions(2).Value), secOther
    WINI "LogDBActions", Cv(chkLogDBActions.Value), secMain
    WINI "LogCommands", Cv(chkLogAllCommands.Value), secMain
    
    '// Advanced Connection Settings
    WINI "UseBNLS", IIf(cboConnMethod.ListIndex = 0, "Y", "N"), secMain
    WINI "Spoof", cboSpoof.ListIndex, secMain
    WINI "UseProxy", Cv(chkUseProxies.Value), secMain
    WINI "ProxyPort", Trim$(txtProxyPort.Text), secMain
    WINI "ProxyIP", Trim$(txtProxyIP.Text), secMain
    'special case, proxyissocks5 didn't like being set properly
    WINI "ProxyIsSocks5", IIf(optSocks5.Value, "Y", "N"), secMain
    WINI "UDP", Cv(chkUDP.Value), secMain
    
    'Don't save this value until the user has been prompted once, or at least specifically enabled it
    s = ReadCfg(secMain, "UseAltBNLS")
    If ((s = "Y") Or (s = "N") Or (chkBNLSAlt.Value = vbChecked)) Then
        WINI "UseAltBNLS", Cv(chkBNLSAlt.Value), secMain
    End If
    BotVars.UseAltBnls = ((s = "Y") Or (chkBNLSAlt.Value = vbChecked))
    
    '// this section must written _absolutely correctly_ or the SetTimer API call will fail
    s = txtReconDelay.Text
    If LenB(s) > 0 Then
        If StrictIsNumeric(s) Then
            If Val(s) < 60000 Then
                If Val(s) > 0 Then
                    I = Val(s)
                Else
                    I = 1000
                End If
            Else
                I = 60000
            End If
        Else
            I = 1000
        End If
    Else
        I = 1000
    End If
    
    WINI "ReconnectDelay", CStr(I), secMain
    
    '// Interface Settings
    WINI "ChatFont", txtChatFont.Text, secOther
    WINI "ChatSize", txtChatSize.Text, secOther
    WINI "ChanFont", txtChanFont.Text, secOther
    WINI "ChanSize", txtChatSize.Text, secOther
    
    '// Idle message settings
    WINI "Idles", Cv(chkIdles.Value), secMain
    WINI "IdleWait", (Val(txtIdleWait.Text) * 2), secMain
    
    Select Case True
        Case optMsg.Value:       WINI "IdleType", "msg", secMain
        Case optUptime.Value:    WINI "IdleType", "uptime", secMain
        Case optMP3.Value:       WINI "IdleType", "mp3", secMain
        Case optQuote.Value:     WINI "IdleType", "quote", secMain
        Case Else: WINI "IdleType", "msg", secMain
    End Select
    
    WINI "IdleMsg", txtIdleMsg.Text, secMain
    
    '// General Moderation settings
    WINI "Phrasebans", Cv(chkPhrasebans.Value), secOther
    WINI "IPBans", Cv(chkIPBans.Value), secOther
    WINI "BanEvasion", Cv(chkBanEvasion.Value), secOther
    
    Clients(SC) = "STAR"
    Clients(BW) = "SEXP"
    Clients(D2) = "D2DV"
    Clients(D2X) = "D2XP"
    Clients(W3) = "WAR3"
    Clients(W3X) = "W3XP"
    Clients(W2) = "W2BN"

    For I = 0 To 6
        If (chkCBan(I).Value = 1) Then
            If (GetAccess(Clients(I), "GAME").Username = _
                vbNullString) Then
                
                ' redefine array size
                ReDim Preserve DB(UBound(DB) + 1)
                
                With DB(UBound(DB))
                    .Username = Clients(I)
                    .Flags = "B"
                    .ModifiedBy = "(console)"
                    .ModifiedOn = Now
                    .AddedBy = "(console)"
                    .AddedOn = Now
                    .Type = "GAME"
                End With
                
                ' commit modifications
                Call WriteDatabase(GetFilePath("users.txt"))
                
                ' log actions
                If (BotVars.LogDBActions) Then
                    Call LogDBAction(AddEntry, "console", DB(UBound(DB)).Username, "game", _
                        DB(UBound(DB)).Rank, DB(UBound(DB)).Flags)
                End If
            Else
                For j = LBound(DB) To UBound(DB)
                    If ((StrComp(DB(j).Username, Clients(I), vbTextCompare) = 0) And _
                        (StrComp(DB(j).Type, "GAME", vbTextCompare) = 0)) Then
                        
                        If (InStr(1, DB(j).Flags, "B", vbBinaryCompare) = 0) Then
                            With DB(j)
                                .Username = Clients(I)
                                .Flags = "B" & .Flags
                                .ModifiedBy = "(console)"
                                .ModifiedOn = Now
                            End With
                            
                            ' log actions
                            If (BotVars.LogDBActions) Then
                                Call LogDBAction(ModEntry, "console", DB(j).Username, "game", _
                                    DB(j).Rank, DB(j).Flags)
                            End If
                            
                            ' commit modifications
                            Call WriteDatabase(GetFilePath("users.txt"))
                            
                            ' break loop
                            Exit For
                        End If
                    End If
                Next j
            End If
        Else
            If (GetAccess(Clients(I), "GAME").Username <> _
                vbNullString) Then

                For j = LBound(DB) To UBound(DB)
                    If ((StrComp(DB(j).Username, Clients(I), vbTextCompare) = 0) And _
                        (StrComp(DB(j).Type, "GAME", vbTextCompare) = 0)) Then
                        
                        If ((Len(DB(j).Flags) > 1) Or _
                            (DB(j).Rank > 0) Or _
                            (Len(DB(j).Groups) > 1)) Then

                            With DB(j)
                                .Username = Clients(I)
                                .Flags = Replace(.Flags, "B", vbNullString)
                                .ModifiedBy = "(console)"
                                .ModifiedOn = Now
                            End With
                            
                            ' log actions
                            If (BotVars.LogDBActions) Then
                                Call LogDBAction(ModEntry, "console", DB(j).Username, "game", _
                                    DB(j).Rank, DB(j).Flags)
                            End If
                            
                            ' commit modifications
                            Call WriteDatabase(GetFilePath("users.txt"))
                        Else
                            Call RemoveItem(Clients(I), "users", _
                                "GAME")
                                
                            ' log actions
                            If (BotVars.LogDBActions) Then
                                Call LogDBAction(RemEntry, "console", DB(j).Username, "game", _
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
    Next I
    
    WINI "QuietTime", Cv(chkQuiet.Value), secMain
    WINI "KickOnYell", Cv(chkKOY.Value), secOther
    WINI "PlugBans", Cv(chkPlugban.Value), secOther
    WINI "Protect", Cv(chkProtect.Value), secMain
    WINI "ProtectMsg", txtProtectMsg.Text, secOther
    WINI "IdleBans", Cv(chkIdlebans.Value), secOther
    WINI "KickIdle", Cv(chkIdleKick.Value), secOther
    WINI "IdleBanDelay", txtIdleBanDelay.Text, secOther
    WINI "PeonBans", chkPeonbans.Value, secOther
    WINI "BanUnderLevel", Val(txtBanW3.Text), secOther
    WINI "BanD2UnderLevel", Val(txtBanD2.Text), secOther
    WINI "LevelBanMsg", txtLevelBanMsg.Text, secOther
    
    '// Greet Message Settings
    WINI "GreetMsg", txtGreetMsg.Text, secOther
    WINI "WhisperGreet", Cv(chkWhisperGreet.Value), secOther
    WINI "UseGreets", Cv(chkGreetMsg.Value), secOther
    
    '// RELOAD CONFIG
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

Sub CAdd(ByVal ColorName As String, ColorValue As Long, Optional Append As Byte)
    mColors(UBound(mColors)) = ColorValue
    
    If Append = 1 Then ColorName = "Event | " & ColorName
    
    cboColorList.AddItem ColorName
    
    ReDim Preserve mColors(UBound(mColors) + 1)
End Sub

Private Sub cmdExport_Click()
    With cDLG
        .filename = vbNullString
        .ShowSave
        If .filename <> vbNullString Then
            SaveColors .filename
            MsgBox "ColorList exported.", vbOKOnly
        End If
    End With
End Sub

Private Sub cmdImport_Click()
    With cDLG
        .filename = vbNullString
        .ShowOpen
        If .filename <> vbNullString Then
            GetColorLists (.filename)
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
            
        If Dir$(GetProfilePath() & "\colors.sclf") <> vbNullString Then
            Kill GetProfilePath() & "\colors.sclf"
            Call GetColorLists
            Call LoadColors
        End If
    End If
End Sub

Private Sub SaveColors(Optional sPath As String)
    Dim f As Integer
    Dim I As Integer
    
    f = FreeFile
    
    If LenB(sPath) = 0 Then sPath = GetProfilePath() & "\colors.sclf"
    
    Open sPath For Random As #f Len = 4
    
    For I = LBound(mColors) To UBound(mColors)
        Put #f, I + 1, CLng(mColors(I))
        'Debug.Print "Putting color; " & i & ":" & mColors(i)
    Next I
    
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
    txtProxyIP.Enabled = chkUseProxies.Value
    txtProxyPort.Enabled = chkUseProxies.Value
    optSocks4.Enabled = chkUseProxies.Value
    optSocks5.Enabled = chkUseProxies.Value
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
    txtExpKey.Enabled = False
    chkUseRealm.Enabled = False
    lblHashPath.Caption = GetGamePath("RATS")
    chkUDP.Enabled = True
End Sub

Sub optWAR3_Click()
    txtExpKey.Enabled = False
    chkUseRealm.Enabled = False
    lblHashPath.Caption = GetGamePath("3RAW")
    chkUDP.Enabled = False
End Sub

Sub optD2DV_Click()
    txtExpKey.Enabled = False
    chkUseRealm.Enabled = True
    lblHashPath.Caption = GetGamePath("VD2D")
    chkUDP.Enabled = False
End Sub

Sub optW2BN_Click()
    txtExpKey.Enabled = False
    chkUseRealm.Enabled = False
    lblHashPath.Caption = GetGamePath("NB2W")
    chkUDP.Enabled = True
End Sub

Sub optSEXP_Click()
    txtExpKey.Enabled = False
    chkUseRealm.Enabled = False
    lblHashPath.Caption = GetGamePath("RATS")
    chkUDP.Enabled = True
End Sub

Sub optD2XP_Click()
    txtExpKey.Enabled = True
    chkUseRealm.Enabled = True
    lblHashPath.Caption = GetGamePath("PX2D")
    chkUDP.Enabled = False
End Sub

Sub optW3XP_Click()
    txtExpKey.Enabled = True
    chkUseRealm.Enabled = False
    lblHashPath.Caption = GetGamePath("PX3W")
    chkUDP.Enabled = False
End Sub

Private Sub chkGreetMsg_Click()
    chkWhisperGreet.Enabled = chkGreetMsg.Value
    txtGreetMsg.Enabled = chkGreetMsg.Value
End Sub

Private Sub chkProtect_Click()
    txtProtectMsg.Enabled = chkProtect.Value
End Sub

'##########################################
' INIT SUBS
'##########################################

Private Sub InitGenMisc()
    chkPAmp.Value = YesToTrue(ReadCfg(OT, "ProfileAmp"), 0)
    chkWhisperCmds.Value = YesToTrue(ReadCfg(MN, "WhisperBack"), 0)
    chkMail.Value = YesToTrue(ReadCfg(OT, "Mail"), 1)
    
    chkDisablePrefix.Value = YesToTrue(ReadCfg(OT, "DisablePrefix"), 0)
    chkDisableSuffix.Value = YesToTrue(ReadCfg(OT, "DisableSuffix"), 0)
    
    'Debug.Print "Loaded value: " & YesToTrue(readcfg(OT, "TTT"), 1) & " converted to " & IIf(YesToTrue(readcfg(OT, "TTT"), 1) = 1, 0, 1)
    
'    chkDisableMonitor.Value = YesToTrue(ReadCFG(OT, "DisableMonitor"), 1)
    chkAllowMP3.Value = YesToTrue(ReadCfg(OT, "AllowMP3"), 1)
    chkConnectOnStartup.Value = YesToTrue(ReadCfg(MN, "ConnectOnStartup"), 0)
    chkMinimizeOnStartup.Value = YesToTrue(ReadCfg(MN, "MinimizeOnStartup"), 0)
    chkShowOffline.Value = YesToTrue(ReadCfg(MN, "ShowOfflineFriends"), 0)
    
    chkGameConventions(0).Value = YesToTrue(ReadCfg(OT, "UseGameConventions"), 1)
    Call chkGameConventions_Click(0)
    
    chkGameConventions(1).Value = YesToTrue(ReadCfg(OT, "UseD2GameConventions"), 0)
    chkGameConventions(2).Value = YesToTrue(ReadCfg(OT, "UseW3GameConventions"), 1)
    
    chkURLDetect.Value = YesToTrue(ReadCfg(MN, "URLDetect"), 1)
    chkDoNotUsePacketFList.Value = YesToTrue(ReadCfg(MN, "DoNotUseDirectFList"), 0)
          
    chkBackup.Value = YesToTrue(ReadCfg(MN, "UseBackupChan"), 0)
    Call chkBackup_Click
    
    txtBackupChan.Text = ReadCfg(MN, "BackupChan")
    
    chkLogDBActions.Value = YesToTrue(ReadCfg(MN, "LogDBActions"), 0)
    chkLogAllCommands.Value = YesToTrue(ReadCfg(MN, "LogCommands"), 0)
End Sub

Private Sub InitGenIdles()
    Dim s As String
    
    s = ReadCfg(MN, "IdleWait")
    If StrictIsNumeric(s) Then
        txtIdleWait.Text = Val(s) / 2
    End If
    
    s = ReadCfg(MN, "IdleType")
    Select Case s
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
    
    txtIdleMsg.Text = ReadCfg(MN, "IdleMsg")
    If LenB(txtIdleMsg.Text) = 0 Then txtIdleMsg.Text = "/me is a %v by Stealth - http://www.stealthbot.net"
    
    chkIdles.Value = YesToTrue(ReadCfg(MN, "Idles"), 0)
    Call chkIdles_Click
    
End Sub

Private Sub InitGenGreets()
    txtGreetMsg.Text = ReadCfg(OT, "GreetMsg")
    chkGreetMsg.Value = YesToTrue(ReadCfg(OT, "UseGreets"), 0)
    Call chkGreetMsg_Click
    
    chkWhisperGreet.Value = YesToTrue(ReadCfg(OT, "WhisperGreet"), 0)
End Sub

Private Sub InitBasicConfig()
    Dim s As String
    Dim f As Integer
    
    txtUsername.Text = ReadCfg(MN, "Username")
    txtPassword.Text = ReadCfg(MN, "Password")
    cboCDKey.Text = ReadCfg(MN, "CDKey")
    
    ' Backwards compatibility for old LODKey config entry -a
    txtExpKey.Text = ReadCfg(MN, "ExpKey")
    
    If (txtExpKey.Text = vbNullString) Then
        txtExpKey.Text = ReadCfg(MN, "LODKey")
    End If
    
    txtHomeChan.Text = ReadCfg(MN, "HomeChan")
    txtOwner.Text = ReadCfg(MN, "Owner")
    
    OldBotOwner = txtOwner.Text
    
    s = ReadCfg(MN, "Server")
    
    With cboServer
        .AddItem "useast.battle.net"
        .AddItem "63.240.202.139"
        .AddItem "63.240.202.138"
        .AddItem "63.240.202.134"
        .AddItem "63.240.202.131"
        .AddItem "63.240.202.130"
        .AddItem "63.240.202.129"
        .AddItem "63.240.202.128"
        .AddItem "63.240.202.127"
        .AddItem "63.240.202.126"
        .AddItem "63.240.202.122"
        .AddItem "63.240.202.121"
        .AddItem "63.240.202.120"
        .AddItem ""
        .AddItem "uswest.battle.net"
        .AddItem "63.241.83.7"
        .AddItem "63.241.83.8"
        .AddItem "63.241.83.9"
        .AddItem "63.241.83.11"
        .AddItem "63.241.83.12"
        .AddItem "63.241.83.13"
        .AddItem "63.241.83.107"
        .AddItem "63.241.83.108"
        .AddItem "63.241.83.109"
        .AddItem "63.241.83.110"
        .AddItem "63.241.83.111"
        .AddItem "63.241.83.112"
        .AddItem ""
        .AddItem "europe.battle.net"
        .AddItem "213.248.106.200"
        .AddItem "213.248.106.201"
        .AddItem "213.248.106.202"
        .AddItem "213.248.106.204"
        .AddItem ""
        .AddItem "asia.battle.net"
        .AddItem "211.233.0.49"
        .AddItem "211.233.0.50"
        .AddItem "211.233.0.51"
        .AddItem "211.233.0.52"
        .AddItem "211.233.0.53"
        .AddItem "211.233.0.72"
        .AddItem "211.233.0.73"
        .AddItem "211.233.0.74"
        .AddItem "211.233.0.75"
        .AddItem "211.233.0.76"
        .AddItem "211.233.0.78"
        .AddItem "211.233.0.79"
        .AddItem "211.233.0.80"
        .AddItem ""
        .AddItem s
        
        .Text = s
        
        If LenB(Dir$(GetFilePath("servers.txt"))) > 0 Then
            f = FreeFile
            
            Open GetFilePath("servers.txt") For Input As #f
                
                If LOF(f) > 0 Then
                    Do
                        Line Input #f, s
                        
                        .AddItem s
                    Loop While Not EOF(f)
                End If
                
            Close #f
        End If
        
        
    End With
    
    s = ReadCfg(MN, "Trigger")
    If Len(s) > 1 Then
        s = Mid(s, 2, Len(s) - 2) ' fix it! nou
    Else
        s = "."
    End If
    txtTrigger.Text = s
    
    s = ReadCfg(MN, "Product")
    Select Case StrReverse(UCase(s))
        Case "STAR":    Call optSTAR_Click: optSTAR.Value = True
        Case "SEXP":    Call optSEXP_Click: optSEXP.Value = True
        Case "D2DV":    Call optD2DV_Click: optD2DV.Value = True
        Case "D2XP":    Call optD2XP_Click: optD2XP.Value = True
        Case "W2BN":    Call optW2BN_Click: optW2BN.Value = True
        Case "WAR3":    Call optWAR3_Click: optWAR3.Value = True
        Case "W3XP":    Call optW3XP_Click: optW3XP.Value = True
        Case Else:      Call optSTAR_Click: optSTAR.Value = True
    End Select
    
    chkUseRealm.Value = YesToTrue(ReadCfg(MN, "UseRealm"), 0)
    
    Call LoadCDKeys(cboCDKey)
    
End Sub

Private Sub InitGenMod()
    'Dim s As String
    
    chkPhrasebans.Value = YesToTrue(ReadCfg(OT, "Phrasebans"), 1)
    chkIPBans.Value = YesToTrue(ReadCfg(OT, "IPBans"), 0)
    chkQuiet.Value = YesToTrue(ReadCfg(MN, "QuietTime"), 0)
    chkKOY.Value = YesToTrue(ReadCfg(OT, "KickOnYell"), 0)
    chkPlugban.Value = YesToTrue(ReadCfg(OT, "PlugBans"), 0)
    chkPeonbans.Value = IIf(ReadCfg(OT, "PeonBans") = "1", 1, 0)
    chkBanEvasion.Value = YesToTrue(ReadCfg(OT, "BanEvasion"), 1)
    
    chkProtect.Value = YesToTrue(ReadCfg(MN, "Protect"), 0)
    Call chkProtect_Click
    
    txtProtectMsg.Text = ReadCfg(OT, "ProtectMsg")
    
    chkIdlebans.Value = YesToTrue(ReadCfg(OT, "IdleBans"), 0)
    chkIdleKick.Value = YesToTrue(ReadCfg(OT, "KickIdle"), 0)
    Call chkIdlebans_click
    
    txtIdleBanDelay.Text = ReadCfg(OT, "IdleBanDelay")
    
    ' grab client ban settings from database
    
    If (InStr(1, GetAccess("STAR", "GAME").Flags, "B", _
        vbBinaryCompare) <> 0) Then
    
        chkCBan(SC).Value = 1
    End If
    
    If (InStr(1, GetAccess("SEXP", "GAME").Flags, "B", _
        vbBinaryCompare) <> 0) Then
    
        chkCBan(BW).Value = 1
    End If
    
    If (InStr(1, GetAccess("D2DV", "GAME").Flags, "B", _
        vbBinaryCompare) <> 0) Then
    
        chkCBan(D2).Value = 1
    End If
    
    If (InStr(1, GetAccess("D2XP", "GAME").Flags, "B", _
        vbBinaryCompare) <> 0) Then
    
        chkCBan(D2X).Value = 1
    End If
    
    If (InStr(1, GetAccess("W2BN", "GAME").Flags, "B", _
        vbBinaryCompare) <> 0) Then
    
        chkCBan(W2).Value = 1
    End If
    
    If (InStr(1, GetAccess("WAR3", "GAME").Flags, "B", _
        vbBinaryCompare) <> 0) Then
    
        chkCBan(W3).Value = 1
    End If
    
    If (InStr(1, GetAccess("W3XP", "GAME").Flags, "B", _
        vbBinaryCompare) <> 0) Then
    
        chkCBan(W3X).Value = 1
    End If

    txtLevelBanMsg.Text = ReadCfg(OT, "LevelbanMsg")
    If LenB(txtLevelBanMsg.Text) = 0 Then txtLevelBanMsg.Text = "You are below the required level for entry."
    
    txtBanD2.Text = Val(ReadCfg(OT, "BanD2UnderLevel"))
    txtBanW3.Text = Val(ReadCfg(OT, "BanUnderLevel"))
End Sub

Private Sub InitGenInterface()
    Dim s As String
    
    chkJoinLeaves.Value = YesToTrue(ReadCfg(OT, "JoinLeaves"), 1)
    chkFilter.Value = YesToTrue(ReadCfg(OT, "Filters"), 1)
    chkSplash.Value = YesToTrue(ReadCfg(MN, "ShowSplash"), 0)
    chkUTF8.Value = YesToTrue(ReadCfg(MN, "UTF8"), 1)
    chkFlash.Value = YesToTrue(ReadCfg(OT, "FlashWindow"), 0)
    chkNoTray.Value = YesToTrue(ReadCfg(OT, "NoTray"), 0)
    chkNoAutocomplete.Value = YesToTrue(ReadCfg(OT, "NoAutocomplete"), 0)
    chkNoColoring.Value = YesToTrue(ReadCfg(OT, "NoColoring"), 0)
    
    s = ReadCfg(OT, "Logging")
    If Len(s) > 1 Then s = Left$(s, 1)
    
    If StrictIsNumeric(s) And Val(s) < 5 Then
        If (Val(s) = 2) Then
            cboLogging.ListIndex = 0
        ElseIf (Val(s) = 1) Then
            cboLogging.ListIndex = 1
        ElseIf (Val(s) = 0) Then
            cboLogging.ListIndex = 2
        End If
    End If
    
    s = ReadCfg(OT, "Timestamp")
    If Len(s) > 1 Then s = Left$(s, 1)
    'Debug.Print "Loaded: " & s
    If StrictIsNumeric(s) And Val(s) < 5 Then
        cboTimestamp.ListIndex = Val(s)
    End If
    
    s = ReadCfg(MN, "MaxBacklogSize")
    
    If (s <> vbNullString) Then
        txtMaxBackLogSize.Text = Val(ReadCfg(MN, "MaxBacklogSize"))
    Else
        txtMaxBackLogSize.Text = 10000
    End If
    
    txtMaxLogSize.Text = Val(ReadCfg(MN, "MaxLogFileSize"))
    
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
    
    s = ReadCfg(MN, "UseBNLS")
    If s = "N" Then cboConnMethod.ListIndex = 1 Else cboConnMethod.ListIndex = 0
    
    s = ReadCfg(MN, "Spoof")
    If StrictIsNumeric(s) Then
        If Val(s) < 3 Then
            cboSpoof.ListIndex = Val(s)
        Else
            cboSpoof.ListIndex = 0
        End If
    Else
        cboSpoof.ListIndex = 0
    End If
        
    s = ReadCfg(MN, "UDP")
    If s = "Y" Then chkUDP.Value = 1 Else chkUDP.Value = 0
    
    s = ReadCfg(MN, "UseAltBNLS")
    If s = "Y" Then chkBNLSAlt.Value = 1 Else chkBNLSAlt.Value = 0
    
    txtReconDelay.Text = ReadCfg(MN, "ReconnectDelay")
    If LenB(txtReconDelay.Text) = 0 Then txtReconDelay.Text = 1000
    
    chkUseProxies.Value = YesToTrue(ReadCfg(MN, "UseProxy"), 0)
    Call chkUseProxies_Click
'    Call chkUseProxies_Click
'    If Not YesToTrue(readcfg(MN, "UseProxy"), 0) = 0 Then Call chkUseProxies_Click
    
    txtProxyPort.Text = ReadCfg(MN, "ProxyPort")
    txtProxyIP.Text = ReadCfg(MN, "ProxyIP")
    
    If ReadCfg(MN, "ProxyIsSocks5") = "Y" Then
        optSocks5.Value = True
        optSocks4.Value = False
    Else
        optSocks5.Value = False
        optSocks4.Value = True
    End If
    
    ' Adjust this label 2 pixels down
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
    
    If LenB(Dir$(GetConfigFilePath())) > 0 Then
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

Private Function Cv(ByVal I As Integer) As String
    Select Case I
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
    If (StrComp(InitChanFont, txtChanFont.Text, vbTextCompare)) Then
        WINI "ChanFont", txtChanFont.Text, secOther
        frmChat.lvChannel.Font.Name = txtChanFont.Text
        frmChat.lvClanList.Font.Name = txtChanFont.Text
        frmChat.lvFriendList.Font.Name = txtChanFont.Text
    End If
    
    If (StrComp(InitChatFont, txtChatFont.Text, vbTextCompare)) Then
        WINI "ChatFont", txtChatFont.Text, secOther
        frmChat.rtbChat.Font.Name = txtChatFont.Text
    End If
    
    If Not InitChanSize = CInt(txtChanSize.Text) Then
        WINI "ChanSize", txtChanSize.Text, secOther
        frmChat.lvChannel.Font.Size = CInt(txtChanSize.Text)
        frmChat.lvClanList.Font.Size = CInt(txtChanSize.Text)
        frmChat.lvFriendList.Font.Size = CInt(txtChanSize.Text)
    End If
    
    If Not InitChatSize = CInt(txtChatSize.Text) Then
        WINI "ChatSize", txtChatSize.Text, secOther
        frmChat.rtbChat.Font.Size = CInt(txtChatSize.Text)
    End If
End Sub

Sub LoadColors()
    ReDim mColors(0)
    cboColorList.Clear
    
    With FormColors
        CAdd "Current Channel Label | Background", .ChannelLabelBack
        CAdd "Current Channel Label | Text", .ChannelLabelText
        CAdd "Channel List | Background", .ChannelListBack
        CAdd "Channel List | Text", .ChannelListText
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
