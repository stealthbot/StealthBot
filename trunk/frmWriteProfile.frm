VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmProfile 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Profile Viewer"
   ClientHeight    =   4455
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   7905
   Icon            =   "frmWriteProfile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   7905
   StartUpPosition =   1  'CenterOwner
   Begin RichTextLib.RichTextBox rtbLocation 
      Height          =   285
      Left            =   1200
      TabIndex        =   8
      Top             =   1200
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   503
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      TextRTF         =   $"frmWriteProfile.frx":0CCA
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   3840
      Width           =   855
   End
   Begin RichTextLib.RichTextBox rtbProfile 
      Height          =   2775
      Left            =   1200
      TabIndex        =   6
      Top             =   1560
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   4895
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmWriteProfile.frx":0D4C
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
   Begin RichTextLib.RichTextBox rtbSex 
      Height          =   285
      Left            =   1200
      TabIndex        =   9
      Top             =   840
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   503
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      TextRTF         =   $"frmWriteProfile.frx":0DC7
   End
   Begin RichTextLib.RichTextBox rtbAge 
      Height          =   285
      Left            =   1200
      TabIndex        =   10
      Top             =   480
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   503
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      TextRTF         =   $"frmWriteProfile.frx":0E49
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "Sex"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   855
   End
   Begin VB.Label lblUsername 
      BackColor       =   &H00000000&
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   120
      Width           =   3015
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      X1              =   1080
      X2              =   1080
      Y1              =   840
      Y2              =   5280
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Location"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Age"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   855
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDone_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdDone_Click
    End If
End Sub

Private Sub Form_Load()
    Me.Icon = frmChat.Icon
End Sub

Private Sub Form_Unload(Cancel As Integer)
    rtbLocation.text = vbNullString
    lblUsername.Caption = vbNullString
    rtbAge.text = vbNullString
    rtbProfile.text = vbNullString
    rtbSex.text = vbNullString
    
    cboSendHadFocus = True
End Sub

'RTB ADDCHAT SUBROUTINE - originally written by Grok[vL] - modified to support
'                         logging and timestamps, as well as color decoding.
Sub AddText(ByRef rtb As RichTextBox, ParamArray saElements() As Variant)
    On Error Resume Next
    Dim L As Long
    Dim I As Integer
    
    For I = LBound(saElements) To UBound(saElements) Step 2
        If InStr(1, saElements(I), Chr(0), vbBinaryCompare) > 0 Then _
            KillNull saElements(I)
        
        If Len(saElements(I + 1)) > 0 Then
            With rtb
                .SelStart = Len(.text)
                L = .SelStart
                .SelLength = 0
                .SelColor = saElements(I)
                .SelText = saElements(I + 1) & Left$(vbCrLf, -2 * CLng((I + 1) = UBound(saElements)))
                .SelStart = Len(.text)
            End With
        End If
    Next I
    
    Call ColorModify(rtb, L)
End Sub

Private Sub rtbProfile_KeyPress(KeyAscii As Integer)
    Call Form_KeyPress(KeyAscii)
End Sub

Private Sub txtAge_KeyPress(KeyAscii As Integer)
    Call Form_KeyPress(KeyAscii)
End Sub

Private Sub txtLoc_KeyPress(KeyAscii As Integer)
    Call Form_KeyPress(KeyAscii)
End Sub

Private Sub txtSex_KeyPress(KeyAscii As Integer)
    Call Form_KeyPress(KeyAscii)
End Sub
