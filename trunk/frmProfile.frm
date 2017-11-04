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
      TabIndex        =   10
      Top             =   3720
      Width           =   975
   End
   Begin RichTextLib.RichTextBox rtbField 
      Height          =   285
      Index           =   3
      Left            =   1200
      TabIndex        =   7
      Top             =   1200
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   503
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      AutoVerbMenu    =   -1  'True
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
      TabIndex        =   11
      Top             =   3360
      Width           =   975
   End
   Begin RichTextLib.RichTextBox rtbField 
      Height          =   2775
      Index           =   4
      Left            =   1200
      TabIndex        =   9
      Top             =   1560
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   4895
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmProfile.frx":0D5B
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
   Begin RichTextLib.RichTextBox rtbField 
      Height          =   285
      Index           =   2
      Left            =   1200
      TabIndex        =   5
      Top             =   840
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   503
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmProfile.frx":0DEC
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
   Begin RichTextLib.RichTextBox rtbField 
      Height          =   285
      Index           =   1
      Left            =   1200
      TabIndex        =   3
      Top             =   480
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   503
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmProfile.frx":0E7D
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
   Begin VB.Label lblField 
      BackColor       =   &H00000000&
      Caption         =   "&Sex:"
      BeginProperty Font 
         Name            =   "Tahoma"
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
      Left            =   120
      TabIndex        =   4
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
      TabIndex        =   1
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   3015
   End
   Begin VB.Label lblField 
      BackColor       =   &H00000000&
      Caption         =   "&Description:"
      BeginProperty Font 
         Name            =   "Tahoma"
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
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label lblField 
      BackColor       =   &H00000000&
      Caption         =   "&Location:"
      BeginProperty Font 
         Name            =   "Tahoma"
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
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lblField 
      BackColor       =   &H00000000&
      Caption         =   "&Age:"
      BeginProperty Font 
         Name            =   "Tahoma"
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
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblField 
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

Private Const FIELD_AGE As Integer = 1
Private Const FIELD_SEX As Integer = 2
Private Const FIELD_LOC As Integer = 3
Private Const FIELD_PRF As Integer = 4

Private m_IsWriting As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If m_IsWriting Then
        Call SetProfile(GetRTBText(rtbField(FIELD_LOC)), GetRTBText(rtbField(FIELD_PRF)), GetRTBText(rtbField(FIELD_SEX)))
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = frmChat.Icon
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    lblUsername.Caption = vbNullString
    For i = rtbField.LBound To rtbField.UBound
        rtbField(i).Text = vbNullString
    Next i
    
    cboSendHadFocus = True
End Sub

Public Sub PrepareForProfile(ByVal Username As String, ByVal IsWriting As Boolean)
    ' store for later
    m_IsWriting = IsWriting

    ' correct caps
    If IsWriting Then Username = GetCurrentUsername

    ' set caption
    Caption = StringFormat("{0}Profile - {1}", IIf(IsWriting, "Edit ", vbNullString), Username)

    ' set Username
    lblUsername.Caption = Username

    ' set up command buttons
    cmdCancel.Visible = IsWriting
    cmdOK.Caption = IIf(IsWriting, "&Write", "&Done")

    ' set locked based on mode
    rtbField(FIELD_AGE).Locked = True 'Not IsWriting - always fixed
    rtbField(FIELD_AGE).TabStop = Not IsWriting
    rtbField(FIELD_SEX).Locked = Not IsWriting
    rtbField(FIELD_LOC).Locked = Not IsWriting
    rtbField(FIELD_PRF).Locked = Not IsWriting

    ' if we are writing, request our own profile
    If IsWriting Then
        Call RequestProfile(GetCurrentUsername, reqUserInterface)
    End If
End Sub

Public Sub SetKey(ByVal KeyName As String, ByVal KeyValue As String)
    Dim Index        As Integer
    Dim saElements() As Variant
    Dim arr()        As Variant
    Dim i            As Long
    Dim StyleBold    As Boolean
    Dim StyleItal    As Boolean
    Dim StyleUndl    As Boolean
    Dim StyleStri    As Boolean

    ' make sure shown
    Show

    'frmChat.AddChat vbWhite, "[Profile] " & KeyName & " == " & KeyValue

    Select Case KeyName
        Case "Profile\Age":         Index = 1
        Case "Profile\Sex":         Index = 2
        Case "Profile\Location":    Index = 3
        Case "Profile\Description": Index = 4
        Case Else:                  Exit Sub
    End Select

    With rtbField(Index)
        .Text = vbNullString
        SetTextSelection rtbField(Index), 0, 0
        .SelBold = False
        .SelItalic = False
        .SelUnderline = False
        .SelFontName = rtbField(Index).Font.Name
        .SelColor = vbWhite

        If m_IsWriting Then
            RTBSetSelectedText rtbField(Index), KeyValue
        Else
            ReDim saElements(0 To 2)
            saElements(0) = vbNullString
            saElements(1) = vbWhite
            saElements(2) = KeyValue
            If ApplyGameColors(saElements(), arr()) Then
                saElements() = arr()
            End If
            ' place each element
            For i = LBound(saElements) To UBound(saElements) Step 3
                DisplayRichTextElement rtbField(Index), saElements(), i, StyleBold, StyleItal, StyleUndl, StyleStri
            Next i
        End If
    End With

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
'        If InStr(1, saElements(I), vbNullChar, vbBinaryCompare) > 0 Then _
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

Private Sub rtbField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If (rtbField(Index).Locked) Then
    
        Select Case (KeyCode)
            Case vbKeyReturn
                cmdOK_Click
                
            Case vbKeyEscape
                cmdCancel_Click
        
            Case vbKeyA, vbKeyC, vbKeyX, vbKeyV, vbKeyReturn, vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyHome, vbKeyEnd, vbKeyPageDown, vbKeyPageUp
                ' don't disable these
                
            Case Else
                ' disable CTRL+L, CTRL+E, CTRL+R, CTRL+I and lots of funny ones
                If (Shift = vbCtrlMask) Then KeyCode = 0
        End Select
        
        Exit Sub
    End If
    
    If (rtbField(Index).SelColor <> vbWhite) Then rtbField(Index).SelColor = vbWhite
    
    Select Case (KeyCode)
        Case vbKeyB
            If (Shift = vbCtrlMask) Then
                rtbField(Index).SelText = Chr$(255) & "cb"
            End If
            
        Case vbKeyU
            If (Shift = vbCtrlMask) Then
                rtbField(Index).SelText = Chr$(255) & "cu"
            End If
            
        Case vbKeyI
            If (Shift = vbCtrlMask) Then
                rtbField(Index).SelText = Chr$(255) & "ci"
            End If
            
        Case vbKeyReturn
            If (Shift = vbCtrlMask) Then
                cmdOK_Click
            End If
                
        Case vbKeyEscape
            cmdCancel_Click
            
        Case vbKeyA, vbKeyC, vbKeyX, vbKeyV, vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyHome, vbKeyEnd, vbKeyPageDown, vbKeyPageUp
            ' don't disable these
            
        Case Else
            ' disable CTRL+L, CTRL+E, CTRL+R, CTRL+I and lots of funny ones
            If (Shift = vbCtrlMask) Then KeyCode = 0
            
    End Select
End Sub

Private Sub rtbField_KeyPress(Index As Integer, KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) And Index < rtbField.UBound Then
        KeyAscii = 0
        rtbField(Index + 1).SetFocus
    End If
End Sub
