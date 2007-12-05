VERSION 5.00
Begin VB.Form frmGroupSelection 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Group Selection"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2175
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   2175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton Option4 
      Caption         =   "Clan"
      Enabled         =   0   'False
      Height          =   195
      Left            =   840
      TabIndex        =   6
      Top             =   1080
      Width           =   735
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Game"
      Enabled         =   0   'False
      Height          =   195
      Left            =   840
      TabIndex        =   5
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&OK"
      Height          =   255
      Index           =   0
      Left            =   1175
      TabIndex        =   4
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   255
      Index           =   0
      Left            =   200
      TabIndex        =   3
      Top             =   1680
      Width           =   855
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Dynamic"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Static"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Value           =   -1  'True
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "What kind of group do you wish to create?"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmGroupSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
