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
      Height          =   345
      Index           =   3
      Left            =   1200
      TabIndex        =   7
      Top             =   1200
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   609
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmProfile.frx":0CCA
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
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmProfile.frx":0D45
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
   Begin RichTextLib.RichTextBox rtbField 
      Height          =   345
      Index           =   2
      Left            =   1200
      TabIndex        =   5
      Top             =   840
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   609
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmProfile.frx":0DC0
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
   Begin RichTextLib.RichTextBox rtbField 
      Height          =   345
      Index           =   1
      Left            =   1200
      TabIndex        =   3
      Top             =   480
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   609
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmProfile.frx":0E3B
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
      Height          =   375
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
        Call SetProfile(GetRTBText(rtbField(FIELD_LOC).hWnd), _
                        GetRTBText(rtbField(FIELD_PRF).hWnd), _
                        GetRTBText(rtbField(FIELD_SEX).hWnd))
    End If
    
    Unload Me

End Sub

Private Sub Form_Load()
    Dim i As Integer

    Me.Icon = frmChat.Icon

    #If COMPILE_DEBUG <> 1 Then
        HookWindowProc Me.hWnd
    #End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    
    lblUsername.Caption = vbNullString

    frmChat.cboSendHadFocus = True

    #If COMPILE_DEBUG <> 1 Then
        UnhookWindowProc Me.hWnd
    #End If
End Sub

Public Sub PrepareForProfile(ByVal Username As String, ByVal IsWriting As Boolean)
    Dim i As Integer
    
    ' store for later
    m_IsWriting = IsWriting

    ' correct caps
    If IsWriting Then Username = BotVars.Username

    ' set Username
    With lblUsername
        .Caption = Username
        .Font = Config.ChatFont
        .Font.Size = Config.ChatFontSize
    End With

    ' set caption
    If IsWriting Then
        Caption = StringFormat("Editing Profile - {0}", Username)
    Else
        Caption = StringFormat("Profile - {0}", Username)
    End If

    ' command buttons
    cmdOK.Caption = IIf(IsWriting, "&Write", "&Done")
    cmdOK.Cancel = IsWriting
    cmdCancel.Visible = IsWriting

    ' set up fields
    For i = rtbField.LBound To rtbField.UBound
        With rtbField(i)
            .Font = Config.ChatFont
            .Font.Size = Config.ChatFontSize
            .Text = vbNullString
            .Locked = Not IsWriting Or i = FIELD_AGE
            If .Locked Then
                ' reading
                EnableRichText .hWnd
                EnableURLDetect .hWnd
            Else
                ' writing
                DisableRichText .hWnd
                DisableURLDetect .hWnd
            End If
        End With
    Next i

    ' request own profile
    If IsWriting Then Call RequestProfile(Username, reqUserInterface)

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
        Case "Profile\Age":         Index = FIELD_AGE
        Case "Profile\Sex":         Index = FIELD_SEX
        Case "Profile\Location":    Index = FIELD_LOC
        Case "Profile\Description": Index = FIELD_PRF
        Case Else:                  Exit Sub
    End Select

    With rtbField(Index)
        If Not .Locked Then
            ' writing
            .HideSelection = True
            SetTextSelection .hWnd, 0, -1
            SetSelectedRTBText .hWnd, KeyValue
            SetTextSelection .hWnd, 0, -1
            .SelBold = False
            .SelItalic = False
            .SelUnderline = False
            .SelFontName = .Font.Name
            .SelColor = vbWhite
            SetTextSelection .hWnd, -1, -1
            .HideSelection = False
        Else
            ' reading
            .HideSelection = True
            .Text = vbNullString
            SetTextSelection .hWnd, 0, 0
            .SelBold = False
            .SelItalic = False
            .SelUnderline = False
            .SelFontName = .Font.Name
            .SelColor = vbWhite

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
            SetTextSelection .hWnd, -1, -1
            .HideSelection = False
        End If
    End With

    SetFocus

End Sub

Private Sub rtbField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    With rtbField(Index)
        If Not .Locked Then
            ' writing
            Select Case KeyCode
                Case vbKeyB
                    If Shift = vbCtrlMask Then
                        SetSelectedRTBText .hWnd, Chr$(255) & "cb"
                        KeyCode = 0
                    End If

                Case vbKeyU
                    If Shift = vbCtrlMask Then
                        SetSelectedRTBText .hWnd, Chr$(255) & "cu"
                        KeyCode = 0
                    End If

                Case vbKeyI
                    If Shift = vbCtrlMask Then
                        SetSelectedRTBText .hWnd, Chr$(255) & "ci"
                        KeyCode = 0
                    End If

                Case vbKeyReturn
                    If Shift = vbCtrlMask Then
                        ' CTRL+ENTER: OK
                        KeyCode = 0
                        cmdOK_Click
                    End If

                Case vbKeyC
                    If Shift = vbCtrlMask Then
                        SetClipboardText GetRTBText(.hWnd, True), .hWnd
                        KeyCode = 0
                    End If

                Case vbKeyV
                    If Shift = vbCtrlMask Then
                        Dim sParam As Long, eParam As Long, sLength As Long
                        GetTextSelection .hWnd, sParam, eParam
                        SetSelectedRTBText .hWnd, GetClipboardText(.hWnd, sLength)
                        SetTextSelection .hWnd, 0, -1
                        .SelFontName = .Font.Name
                        SetTextSelection .hWnd, eParam + sLength, eParam + sLength
                        KeyCode = 0
                    End If

                Case vbKeyX
                    If Shift = vbCtrlMask Then
                        SetClipboardText GetRTBText(.hWnd, True), .hWnd
                        SetSelectedRTBText .hWnd, vbNullString
                        KeyCode = 0
                    End If

                Case vbKeyA, vbKeyReturn, vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyHome, vbKeyEnd, vbKeyPageDown, vbKeyPageUp
                    ' don't disable these

                Case Else
                    ' disable CTRL+L, CTRL+E, CTRL+R, CTRL+I and lots of funny ones
                    If (Shift = vbCtrlMask) Then KeyCode = 0

            End Select
        Else
            ' reading
            Select Case KeyCode
                Case vbKeyReturn
                    If Shift = vbCtrlMask Or Index = rtbField.UBound Then
                        ' CTRL+ENTER or ENTER in last field: OK
                        cmdOK_Click
                    End If

                Case vbKeyC, vbKeyX
                    If Shift = vbCtrlMask Then
                        SetClipboardText GetRTBText(.hWnd, True), Me.hWnd
                        KeyCode = 0
                    End If

                Case vbKeyV
                    If Shift = vbCtrlMask Then
                        KeyCode = 0
                    End If

                Case vbKeyA, vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyHome, vbKeyEnd, vbKeyPageDown, vbKeyPageUp
                    ' don't disable these

                Case Else
                    ' disable CTRL+L, CTRL+E, CTRL+R, CTRL+I and lots of funny ones
                    If Shift = vbCtrlMask Then KeyCode = 0

            End Select
        End If
    End With

End Sub

Private Sub rtbField_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And Index < rtbField.UBound Then
        KeyAscii = 0
        rtbField(Index + 1).SetFocus
    End If
End Sub
