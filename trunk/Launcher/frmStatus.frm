VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmStatus 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Launcher Status"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6525
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
   ScaleHeight     =   2475
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox rtbStatus 
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   2990
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      AutoVerbMenu    =   -1  'True
      OLEDropMode     =   0
      TextRTF         =   $"frmStatus.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const OBJECT_NAME As String = "frmStatus"

Private m_oldhwnd As Long

Private Sub Form_Load()
On Error GoTo ERROR_HANDLER:
    Me.Icon = frmLauncher.Icon
    
    HookWindowProc Me.hWnd
    EnableURLDetection rtbStatus.hWnd

    Exit Sub
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "Form_Load"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ERROR_HANDLER:
    
    Cancel = IIf(bIsClosing, 0, 1)
    Me.Hide

    Exit Sub
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "Form_QueryUnload"
End Sub

Private Sub Form_Resize()
On Error GoTo ERROR_HANDLER:

    rtbStatus.Width = Me.ScaleWidth
    rtbStatus.Height = Me.ScaleHeight
    
    Exit Sub
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "Form_Reload"
End Sub
