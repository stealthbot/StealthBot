VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmWhisperWindow 
   BackColor       =   &H00000000&
   Caption         =   "< account name >"
   ClientHeight    =   3270
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   ScaleHeight     =   3270
   ScaleWidth      =   4695
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog cdl 
      Left            =   2160
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtSend 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   4455
   End
   Begin RichTextLib.RichTextBox rtbWhispers 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   4683
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmWhisperWindow.frx":0000
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
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuSave 
         Caption         =   "&Save Conversation"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuIgnoreAndClose 
         Caption         =   "&Ignore and Close"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHide 
         Caption         =   "&Hide"
      End
   End
End
Attribute VB_Name = "frmWhisperWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_sWhisperTo As String
Private m_imyIndex As Integer
Private m_StartDate As Date
Public Shown As Boolean

'Public MyOldWndProc As Long

Public Property Get StartDate() As Date
    m_StartDate = m_StartDate
End Property

Public Property Let StartDate(ByVal sNewStartDate As Date)
    m_StartDate = sNewStartDate
End Property

Public Property Get sWhisperTo() As String
    sWhisperTo = m_sWhisperTo
End Property

Public Property Let sWhisperTo(ByVal ssWhisperTo As String)
    If InStr(ssWhisperTo, "*") Then
        ssWhisperTo = Mid$(ssWhisperTo, InStr(ssWhisperTo, "*") + 1)
    End If
    
    m_sWhisperTo = ssWhisperTo
    
    Me.Caption = "Whisper Window: " & ssWhisperTo
End Property

Public Property Get myIndex() As Integer
    myIndex = m_imyIndex
End Property

Public Property Let myIndex(ByVal imyIndex As Integer)
    m_imyIndex = imyIndex
End Property

Private Sub Form_Load()
    Me.Icon = frmChat.Icon
    
    With frmChat.rtbChat
        rtbWhispers.Font.Name = .Font.Name
        rtbWhispers.Font.Bold = .Font.Bold
        rtbWhispers.Font.Size = .Font.Size
        txtSend.Font.Name = .Font.Name
        txtSend.Font.Bold = .Font.Bold
        txtSend.Font.Size = .Font.Size
    End With
    
    Form_Resize
    
'    If Me.MyOldWndProc = 0 Then
'        Me.MyOldWndProc = SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf WWNewWndProc)
'    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call DestroyWW(m_imyIndex)
End Sub

Private Sub mnuClose_Click()
    Call DestroyWW(m_imyIndex)
End Sub

Private Sub mnuHide_Click()
    Shown = False
    Me.Hide
End Sub

Private Sub mnuIgnoreAndClose_Click()
    frmChat.AddQ "/ignore " & m_sWhisperTo
    Call DestroyWW(m_imyIndex)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Dim SPACER As Long
    SPACER = rtbWhispers.Left
    
    With rtbWhispers
        .Height = Me.ScaleHeight - txtSend.Height - SPACER - 100
        .Width = Me.ScaleWidth - (SPACER * 2)
        .Font.Name = frmChat.rtbChat.Font.Name
        .Font.Size = frmChat.rtbChat.Font.Size
        .BackColor = frmChat.rtbChat.BackColor
    End With
    
    With txtSend
        .Width = rtbWhispers.Width
        .Top = rtbWhispers.Top + rtbWhispers.Height + 10
        .Font.Name = frmChat.cboSend.Font.Name
        .Font.Size = frmChat.cboSend.Font.Size
        .BackColor = frmChat.cboSend.BackColor
    End With
    
    txtSend.SetFocus
End Sub

Private Sub mnuSave_Click()
    Dim ToSave() As String
    Dim f As Integer, i As Integer
    Dim tUsername As String, tMessage As String

    With cdl
        .InitDir = CurDir$()
        .Filter = ".htm|HTML Documents"
        .ShowSave
    
        If LenB(.FileName) > 0 Then
            ToSave() = Split(rtbWhispers.Text, vbCrLf)
            f = FreeFile
            
            If InStr(1, .FileName, ".") = 0 Then
                .FileName = .FileName & ".htm"
            End If
            
            Open .FileName For Output As #f
                Print #f, "<html><head>"
                Print #f, "<title>StealthBot Conversation Log: " & GetCurrentUsername & " and " & m_sWhisperTo & "</title></head>"
                Print #f, "<body bgcolor='#000000'>"
                
                Print #f, "<p><font color='#FFFFFF'><b>"
                Print #f, "StealthBot Conversation Log, between " & GetCurrentUsername & " and " & m_sWhisperTo & ".<br />"
                Print #f, "Conversation began: " & Format(m_StartDate, "HH:MM:SS, m/dd/yyyy")
                Print #f, "</b></font></p>"
                
                Print #f, "<p>"
                
                For i = 0 To UBound(ToSave)
                    If LenB(ToSave(i)) > 0 Then
                        If InStr(ToSave(i), ":") > 0 Then
                            tMessage = Mid$(ToSave(i), InStr(ToSave(i), ":") + 2)
                            tUsername = Split(ToSave(i), " ")(1)
                            tUsername = Left$(tUsername, InStr(tUsername, ":") - 1)
                        Else
                            tMessage = ToSave(i)
                        End If

                        If StrComp(tUsername, GetCurrentUsername, vbTextCompare) = 0 Then
                            Print #f, "<font size='-1' color='#" & VBHexToHTMLHex(Hex(RTBColors.TalkBotUsername)) & "'><b>";
                        Else
                            Print #f, "<font size='-1' color='#" & VBHexToHTMLHex(Hex(RTBColors.WhisperUsernames)) & "'><b>";
                        End If
                            Print #f, "» " & tUsername & "</b></font>"
                        
                        Print #f, "<font size='-1' color='#" & VBHexToHTMLHex(Hex(RTBColors.WhisperCarats)) & "'><b>";
                            Print #f, ":</b></font> "
                            
                        Print #f, "<font size='-1' color='#" & VBHexToHTMLHex(Hex(RTBColors.WhisperText)) & "'>";
                            Print #f, tMessage & "</font><br />"
                              
                    End If
                Next i
                
                Print #f, "</p>"
                Print #f, "</body></html>"
            Close #f
            
            AddWhisper vbGreen, "» Conversation saved."
        End If
    End With
End Sub

Private Sub rtbWhispers_KeyDown(KeyCode As Integer, Shift As Integer)
    'Disable Ctrl+L, Ctrl+E, and Ctrl+R
    If (Shift = vbCtrlMask) And ((KeyCode = vbKeyL) Or (KeyCode = vbKeyE) Or (KeyCode = vbKeyR)) Then
        KeyCode = 0
    End If
End Sub

Private Sub rtbWhispers_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 32) Then
        Exit Sub
    End If

    txtSend.SetFocus
    
    txtSend.SelText = Chr$(KeyAscii)
End Sub

Private Sub txtSend_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        frmChat.AddQ "/w " & IIf(Dii, "*", "") & m_sWhisperTo & Space(1) & txtSend.Text
        KeyAscii = 0
        txtSend.Text = ""
    End If
    
    Dim x() As String
    Dim i As Integer
    
    If KeyAscii = 22 Then
        On Error Resume Next
        
        If InStr(1, Clipboard.GetText, Chr(13), vbTextCompare) <> 0 Then
        
            x() = Split(Clipboard.GetText, Chr(10))
            If UBound(x) > 0 Then
                For i = LBound(x) To UBound(x)
                    If i = LBound(x) Then x(i) = txtSend.Text & x(i)
                
                    x(i) = Replace(x(i), Chr(13), vbNullString)
                    
                    If x(i) <> vbNullString Then
                        frmChat.AddQ "/w " & m_sWhisperTo & Space(1) & x(i)
                    End If
                Next i
                txtSend.Text = vbNullString
                KeyAscii = 0
            End If
        End If
    End If
End Sub

Sub AddWhisper(ParamArray saElements() As Variant)

    On Error Resume Next
    Dim s As String
    Dim L As Long
    Dim i As Integer, oldSelStart As Integer, oldSelLength As Integer
    
    oldSelStart = txtSend.selStart
    oldSelStart = oldSelStart + txtSend.selLength
    
    If GetForegroundWindow() = Me.hWnd Then
        rtbWhispers.Locked = True
    End If
    
    If Not BotVars.LockChat Then
        With rtbWhispers
            .selStart = Len(.Text)
            .selLength = 0
            .SelColor = RTBColors.TimeStamps
            .SelText = s
            .selStart = Len(.Text)
        End With
        
        For i = LBound(saElements) To UBound(saElements) Step 2
            If InStr(1, saElements(i), Chr(0), vbBinaryCompare) > 0 Then _
                KillNull saElements(i)
            
            If Len(saElements(i + 1)) > 0 Then
                With rtbWhispers
                    .selStart = Len(.Text)
                    L = .selStart
                    .selLength = 0
                    .SelColor = saElements(i)
                    .SelText = saElements(i + 1) & Left$(vbCrLf, -2 * CLng((i + 1) = UBound(saElements)))
                    .selStart = Len(.Text)
                End With
            End If
        Next i
        
        Call ColorModify(rtbWhispers, L)
        
        txtSend.selStart = oldSelStart
        txtSend.selLength = oldSelLength
    End If
    
'    If rtbWhispers.Locked Then
'        rtbWhispers.Locked = False
'    End If
End Sub
