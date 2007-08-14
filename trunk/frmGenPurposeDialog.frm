VERSION 5.00
Begin VB.Form frmGenPurposeDialog 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "<caption>"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraMain 
      BackColor       =   &H00000000&
      ForeColor       =   &H00000000&
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      Begin VB.CommandButton Command1 
         Caption         =   "&No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   2
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton cmdYes 
         Caption         =   "&Yes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   1
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "<caption>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   3975
      End
   End
End
Attribute VB_Name = "frmGenPurposeDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' StealthBot - frmGeneralPurposeDialog
'   January 17th, 2006
'   General-purpose dialog box - yes? no!

Option Explicit

Private m_Option As Integer


