VERSION 5.00
Begin VB.Form frmDBType 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Database Type"
   ClientHeight    =   900
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2160
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
   ScaleHeight     =   900
   ScaleWidth      =   2160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnSave 
      Caption         =   "OK"
      Height          =   300
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.ComboBox cbxChoice 
      BackColor       =   &H00993300&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmDBType.frx":0000
      Left            =   120
      List            =   "frmDBType.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton btnDelete 
      Caption         =   "Cancel"
      Height          =   300
      Left            =   1080
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "frmDBType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_path As String

Private Sub Form_Load()
    cbxChoice.ListIndex = 0
End Sub

Private Sub btnDelete_Click()
    Unload Me
End Sub

Private Sub btnSave_Click(Index As Integer)
    Call frmDBManager.ImportDatabase(m_path, cbxChoice.ListIndex)
    
    Unload Me
End Sub

Public Sub setFilePath(strPath As String)
    m_path = strPath
End Sub
