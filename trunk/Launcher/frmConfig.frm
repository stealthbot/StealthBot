VERSION 5.00
Begin VB.Form frmConfig 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create Profile"
   ClientHeight    =   765
   ClientLeft      =   105
   ClientTop       =   495
   ClientWidth     =   4185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   765
   ScaleWidth      =   4185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkAutoClose 
      BackColor       =   &H00000000&
      Caption         =   "Automatically close launcher when profile is loaded."
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   3975
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
      Left            =   2640
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
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
      Left            =   1200
      TabIndex        =   0
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const OBJECT_NAME As String = "frmConfig"

Private Sub Form_Load()
On Error GoTo ERROR_HANDLER
    Me.Caption = "Launcher Settings"
    Me.Icon = frmLauncher.Icon
    
    If (cConfig Is Nothing) Then Set cConfig = New clsConfig
    
    chkAutoClose.Value = IIf(cConfig.AutoClose, 1, 0)
    
    Exit Sub
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "Form_Load"
End Sub

Private Sub cmdApply_Click()
On Error GoTo ERROR_HANDLER
    
    If (cConfig Is Nothing) Then Set cConfig = New clsConfig
    
    With cConfig
        .AutoClose = (chkAutoClose.Value = 1)
        
        .SaveConfig
    End With
    
    Unload Me
    
    Exit Sub
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "cmdApply_Click"
End Sub

Private Sub cmdCancel_Click()
On Error GoTo ERROR_HANDLER

    Set cConfig = Nothing
    Set cConfig = New clsConfig
    
    Unload Me
    
    Exit Sub
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "cmdCancel_Click"
End Sub
