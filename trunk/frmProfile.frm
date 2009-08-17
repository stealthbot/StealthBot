VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmProfile 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Profile Viewer"
   ClientHeight    =   4455
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   7905
   Icon            =   "frmProfile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   7905
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Write"
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
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   3720
      Width           =   855
   End
   Begin RichTextLib.RichTextBox rtbLocation 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   1200
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   503
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      TextRTF         =   $"frmProfile.frx":0CCA
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
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3360
      Width           =   855
   End
   Begin RichTextLib.RichTextBox rtbProfile 
      Height          =   2775
      Left            =   1200
      TabIndex        =   2
      Top             =   1560
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   4895
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmProfile.frx":0D45
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
      TabIndex        =   0
      Top             =   840
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   503
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      TextRTF         =   $"frmProfile.frx":0DC0
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
   Begin RichTextLib.RichTextBox rtbAge 
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      Top             =   480
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   503
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      TextRTF         =   $"frmProfile.frx":0E3B
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
      TabIndex        =   9
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
      TabIndex        =   6
      Top             =   120
      Width           =   3015
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      X1              =   1080
      X2              =   1080
      Y1              =   120
      Y2              =   4320
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
      TabIndex        =   11
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
      TabIndex        =   10
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
      TabIndex        =   8
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
      TabIndex        =   7
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

Private m_IsWriting As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If m_IsWriting Then SetProfile rtbLocation.Text, rtbProfile.Text, rtbSex.Text
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = frmChat.Icon
End Sub

Private Sub Form_Unload(Cancel As Integer)
    lblUsername.Caption = vbNullString
    rtbAge.Text = vbNullString
    rtbSex.Text = vbNullString
    rtbLocation.Text = vbNullString
    rtbProfile.Text = vbNullString
    
    cboSendHadFocus = True
End Sub

Public Sub PrepareForProfile(ByVal Username As String, ByVal IsWriting As Boolean)
    ' store for later
    m_IsWriting = IsWriting
    
    ' set caption
    Caption = IIf(IsWriting, "Profile Writer - " & GetCurrentUsername, "Profile Viewer - " & Username)
    
    ' set Username
    lblUsername.Caption = IIf(IsWriting, GetCurrentUsername, Username)
    
    ' set up command buttons
    cmdCancel.Visible = IsWriting
    cmdOK.Caption = IIf(IsWriting, "&Write", "&Done")
    
    ' set locked based on mode
    rtbAge.Locked = True 'Not IsWriting - always fixed
    rtbSex.Locked = Not IsWriting
    rtbLocation.Locked = Not IsWriting
    rtbProfile.Locked = Not IsWriting
    
    ' if we are writing, request our own profile
    If IsWriting Then
        ProfileRequest = True
        RequestProfile GetCurrentUsername
    End If
End Sub

Public Sub SetKey(ByVal KeyName As String, ByVal KeyValue As String)
    Dim rtb As RichTextBox
    
    ' make sure shown
    Show
    
    'frmChat.AddChat vbWhite, "[Profile] " & KeyName & " == " & KeyValue
    
    Select Case KeyName
        Case "Profile\Age"
            Set rtb = rtbAge
        Case "Profile\Location"
            Set rtb = rtbLocation
        Case "Profile\Description"
            Set rtb = rtbProfile
        Case "Profile\Sex"
            Set rtb = rtbSex
        Case Else
            Exit Sub
    End Select
    
    rtb.Text = vbNullString
    
    rtb.selStart = 0
    rtb.selLength = 0
    rtb.SelColor = vbWhite
    rtb.SelText = KeyValue
    
    If m_IsWriting = False Then Call ColorModify(rtb, 0)
    
    SetFocus
End Sub

'RTB ADDCHAT SUBROUTINE - originally written by Grok[vL] - modified to support
'                         logging and timestamps, as well as color decoding.
'Sub AddText(ByRef rtb As RichTextBox, ParamArray saElements() As Variant)
'    On Error Resume Next
'    Dim L As Long
'    Dim I As Integer
'
'    For I = LBound(saElements) To UBound(saElements) Step 2
'        If InStr(1, saElements(I), Chr(0), vbBinaryCompare) > 0 Then _
'            KillNull saElements(I)
'
'        If Len(saElements(I + 1)) > 0 Then
'            With rtb
'                .selStart = Len(.Text)
'                L = .selStart
'                .selLength = 0
'                .SelColor = saElements(I)
'                .SelText = saElements(I + 1) & Left$(vbCrLf, -2 * CLng((I + 1) = UBound(saElements)))
'                .selStart = Len(.Text)
'            End With
'        End If
'    Next I
'
'    Call ColorModify(rtb, L)
'End Sub
