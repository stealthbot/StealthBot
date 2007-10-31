VERSION 5.00
Begin VB.Form frmScriptUI 
   BackColor       =   &H00000000&
   Caption         =   "Scripting UI"
   ClientHeight    =   3195
   ClientLeft      =   450
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmr 
      Index           =   0
      Left            =   1200
      Top             =   0
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   0
      Left            =   840
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmd 
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lbl 
      BackColor       =   &H00000000&
      Caption         =   "lbl"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "frmScriptUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'10-29-07 - Hdx - Allow users to create a UI form from scripts

Private strPrefix As String
Private strFormName As String
Private oControls As Object

'Public Enum ObjectProps
'    btWidth = 0         'Must be an Int
'    btHeight = 1        'Must be an Int
'    btLeft = 2          'Must be an Int
'    btTop = 3           'Must be an Int
'    btVisable = 4       'Must be a boolean: 0, 1, true, false, "true", "false"
'    btBackColor = 5     'Must be an Int
'    btCaption = 6       'Can be anything
'    btWindowState = 7   'Must be One of the following: 0: Normal, 1: Minimized, 2: Maximized
'    btBorderStyle = 8   'Must be one of the following: 0: None, 1: FixedSingle, 2: Sizable, 3: FixedDialog, 4: Fixed Tool, 5: Sizable Tool
'    btToolTip = 9       'Can be anything
'    btText = 10         'Can be anything
'    btTag = 11          'Can be anything
'    btScrollBars = 12   'Must be one of the following: 0: None, 1: Horizontal, 2: Vertical, 3: Both
'    btRightToLeft = 13  'Must be a boolean: 0, 1, true, false, "true", "false"
'    btPasswordChar = 14 'Must be a charecter (Is trimmed to 1 chr if string)
'    btMultiLine = 15    'Must be a boolean: 0, 1, true, false, "true", "false"
'    btMaxLength = 16    'Must be an Int
'    btLocked = 17       'Must be a boolean: 0, 1, true, false, "true", "false"
'    btForeColor = 18    'Must be an Int
'    btEnabled = 19      'Must be a boolean: 0, 1, true, false, "true", "false"
'    btBackStyle = 20    'Must be one of the following: 0: Transparent, 1: Opaque
'    btInterval = 21     'Must be an Int
'End Enum

Public Property Let Prefix(strPre As String)
  strPrefix = strPre
End Property
Public Property Get Prefix() As String
  Prefix = strPrefix
End Property
Public Property Let FormName(strName As String)
  strFormName = strName
End Property
Public Property Get FormName() As String
  FormName = strFormName
End Property

Public Function AddCommandButton(ByVal strName As String) As Boolean
  If (oControls.Exists(strName)) Then
    AddCommandButton = False
  Else
    Dim Index As Integer
    Index = cmd.UBound + 1
    Load cmd(Index)
    cmd(Index).Visible = True
    cmd(Index).Tag = strName
    oControls.Add strName, cmd(Index)
    Form_Resize
    AddCommandButton = True
  End If
End Function

Public Function AddLabel(ByVal strName As String) As Boolean
  If (oControls.Exists(strName)) Then
    AddLabel = False
  Else
    Dim Index As Integer
    Index = lbl.UBound + 1
    Load lbl(Index)
    lbl(Index).Visible = True
    lbl(Index).Tag = strName
    oControls.Add strName, cmd(Index)
    Form_Resize
    AddLabel = True
  End If
End Function

Public Function AddTextBox(ByVal strName As String) As Boolean
  If (oControls.Exists(strName)) Then
    AddTextBox = False
  Else
    Dim Index As Integer
    Index = txt.UBound + 1
    Load txt(Index)
    txt(Index).Visible = True
    txt(Index).Tag = strName
    oControls.Add strName, txt(Index)
    Form_Resize
    AddTextBox = True
  End If
End Function

Public Function AddTimer(ByVal strName As String) As Boolean
  If (oControls.Exists(strName)) Then
    AddTimer = False
  Else
    Dim Index As Integer
    Index = txt.UBound + 1
    Load tmr(Index)
    tmr(Index).Tag = strName
    oControls.Add strName, tmr(Index)
    Form_Resize
    AddTimer = True
  End If
End Function

Public Sub SetObjectProperty(ByVal strObject As String, ByVal btProperty As ObjectProps, ByVal vValue As Variant)
  If (oControls.Exists(strObject)) Then
    Dim obj As Object
    Set obj = oControls.Item(strObject)
    Select Case btProperty
      Case ObjectProps.btWidth:        obj.Width = CInt(vValue)
      Case ObjectProps.btHeight:       obj.Height = CInt(vValue)
      Case ObjectProps.btTop:          obj.Top = CInt(vValue)
      Case ObjectProps.btLeft:         obj.Left = CInt(vValue)
      Case ObjectProps.btVisable:      obj.visable = CBool(vValue)
      Case ObjectProps.btBackColor:    obj.BackColor = CInt(vValue)
      Case ObjectProps.btCaption:      obj.Caption = vValue
      Case ObjectProps.btBorderStyle:  obj.BorderStyle = CInt(vValue)
      Case ObjectProps.btToolTip:      obj.ToolTipText = vValue
      Case ObjectProps.btText:         obj.text = vValue
      Case ObjectProps.btScrollBars:   obj.ScrollBars = CInt(vValue)
      Case ObjectProps.btRightToLeft:  obj.RightToLeft = CBool(vValue)
      Case ObjectProps.btPasswordChar: obj.PasswordChar = Left(vValue, 1)
      Case ObjectProps.btMultiLine:    obj.MultiLine = CBool(vValue)
      Case ObjectProps.btMaxLength:    obj.MaxLength = CInt(vValue)
      Case ObjectProps.btLocked:       obj.Locked = CBool(vValue)
      Case ObjectProps.btForeColor:    obj.ForeColor = CInt(vValue)
      Case ObjectProps.btEnabled:      obj.Enabled = CBool(vValue)
      Case ObjectProps.btBackStyle:    obj.BackStyle = CInt(vValue)
      Case ObjectProps.btInterval:     obj.Interval = CInt(vValue)
    End Select
    Set obj = Nothing
  End If
End Sub

Public Function GetObjectProperty(ByVal strObject As String, ByVal btProperty As ObjectProps)
  If (oControls.Exists(strObject)) Then
    Dim obj As Object
    Set obj = oControls.Item(strObject)
    Select Case btProperty
      Case ObjectProps.btWidth:        GetObjectProperty = obj.Width
      Case ObjectProps.btHeight:       GetObjectProperty = obj.Height
      Case ObjectProps.btTop:          GetObjectProperty = obj.Top
      Case ObjectProps.btLeft:         GetObjectProperty = obj.Left
      Case ObjectProps.btVisable:      GetObjectProperty = obj.visable
      Case ObjectProps.btBackColor:    GetObjectProperty = obj.BackColor
      Case ObjectProps.btCaption:      GetObjectProperty = obj.Caption
      Case ObjectProps.btBorderStyle:  GetObjectProperty = obj.BorderStyle
      Case ObjectProps.btToolTip:      GetObjectProperty = obj.ToolTipText
      Case ObjectProps.btText:         GetObjectProperty = obj.text
      Case ObjectProps.btScrollBars:   GetObjectProperty = obj.ScrollBars
      Case ObjectProps.btRightToLeft:  GetObjectProperty = obj.RightToLeft
      Case ObjectProps.btPasswordChar: GetObjectProperty = obj.PasswordChar
      Case ObjectProps.btMultiLine:    GetObjectProperty = obj.MultiLine
      Case ObjectProps.btMaxLength:    GetObjectProperty = obj.MaxLength
      Case ObjectProps.btLocked:       GetObjectProperty = obj.Locked
      Case ObjectProps.btForeColor:    GetObjectProperty = obj.ForeColor
      Case ObjectProps.btEnabled:      GetObjectProperty = obj.Enabled
      Case ObjectProps.btBackStyle:    GetObjectProperty = obj.BackStyle
      Case ObjectProps.btInterval:     GetObjectProperty = obj.Interval
    End Select
    Set obj = Nothing
  End If
End Function

'//////////////////////////////////////////////////////
'//Events
'//////////////////////////////////////////////////////

Private Sub Form_Load()
  Me.Icon = frmChat.Icon
  Set oControls = CreateObject("Scripting.Dictionary")
  On Error Resume Next
  frmChat.SControl.Run strPrefix & "_" & strFormName & "_Load"
End Sub

Private Sub Form_Activate()
  On Error Resume Next
  frmChat.SControl.Run strPrefix & "_" & strFormName & "_Activate"
End Sub

Private Sub Form_Click()
  On Error Resume Next
  frmChat.SControl.Run strPrefix & "_" & strFormName & "_Click"
End Sub

Private Sub Form_DblClick()
  On Error Resume Next
  frmChat.SControl.Run strPrefix & "_" & strFormName & "_DblClick"
End Sub

Private Sub Form_Deactivate()
  On Error Resume Next
  frmChat.SControl.Run strPrefix & "_" & strFormName & "_Deactivate"
End Sub

Private Sub Form_GotFocus()
  On Error Resume Next
  frmChat.SControl.Run strPrefix & "_" & strFormName & "_GotFocus"
End Sub

Private Sub Form_Initialize()
  On Error Resume Next
  frmChat.SControl.Run strPrefix & "_" & strFormName & "_Initialize"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error Resume Next
  frmChat.SControl.Run strPrefix & "_" & strFormName & "_KeyDown", KeyCode, Shift
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  On Error Resume Next
  frmChat.SControl.Run strPrefix & "_" & strFormName & "_KeyPress", KeyAscii
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  On Error Resume Next
  frmChat.SControl.Run strPrefix & "_" & strFormName & "_KeyUp", KeyCode, Shift
End Sub

Private Sub Form_LostFocus()
  On Error Resume Next
  frmChat.SControl.Run strPrefix & "_" & strFormName & "_LostFocus"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
  frmChat.SControl.Run strPrefix & "_" & strFormName & "_MouseDown", Button, Shift, X, Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
  frmChat.SControl.Run strPrefix & "_" & strFormName & "_MouseMove", Button, Shift, X, Y
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
  frmChat.SControl.Run strPrefix & "_" & strFormName & "_MouseUp", Button, Shift, X, Y
End Sub

Private Sub Form_Paint()
  On Error Resume Next
  frmChat.SControl.Run strPrefix & "_" & strFormName & "_Paint"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error Resume Next
  frmChat.SControl.Run strPrefix & "_" & strFormName & "_LostFocus"
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  frmChat.SControl.Run strPrefix & "_" & strFormName & "_Resize"
End Sub

Private Sub Form_Terminate()
  On Error Resume Next
  frmChat.SControl.Run strPrefix & "_" & strFormName & "_Terminate"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set oControls = Nothing
  On Error Resume Next
  frmChat.SControl.Run strPrefix & "_" & strFormName & "_UnLoad", Cancel
  Debug.Print Cancel
End Sub

Private Sub cmd_LostFocus(Index As Integer)
  On Error Resume Next
  frmChat.SControl.Run strPrefix & "_" & strFormName & "_" & cmd(Index).Tag & "_LostFocus"
End Sub

Private Sub cmd_GotFocus(Index As Integer)
  On Error Resume Next
  frmChat.SControl.Run strPrefix & "_" & strFormName & "_" & cmd(Index).Tag & "_GotFocus"
End Sub

Private Sub cmd_KeyPress(Index As Integer, KeyAscii As Integer)
  On Error Resume Next
  frmChat.SControl.Run strPrefix & "_" & strFormName & "_" & cmd(Index).Tag & "_KeyPress", KeyAscii
End Sub

Private Sub cmd_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
  On Error Resume Next
  frmChat.SControl.Run strPrefix & "_" & strFormName & "_" & cmd(Index).Tag & "_KeyUp", KeyCode, Shift
End Sub

Private Sub cmd_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
  frmChat.SControl.Run strPrefix & "_" & strFormName & "_" & cmd(Index).Tag & "_MouseUp", Button, Shift, X, Y
End Sub

Private Sub cmd_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
  frmChat.SControl.Run strPrefix & "_" & strFormName & "_" & cmd(Index).Tag & "_MouseUp", Button, Shift, X, Y
End Sub

Private Sub cmd_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
  frmChat.SControl.Run strPrefix & "_" & strFormName & "_" & cmd(Index).Tag & "_MouseUp", Button, Shift, X, Y
End Sub

Private Sub cmd_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  On Error Resume Next
  frmChat.SControl.Run strPrefix & "_" & strFormName & "_" & cmd(Index).Tag & "_KeyDown", KeyCode, Shift
End Sub

Private Sub cmd_Click(Index As Integer)
  On Error Resume Next
  frmChat.SControl.Run strPrefix & "_" & strFormName & "_" & cmd(Index).Tag & "_Click"
End Sub

Private Sub lbl_Change(Index As Integer)
  On Error Resume Next
  frmChat.SControl.Run strPrefix & "_" & strFormName & "_" & lbl(Index).Tag & "_Change"
End Sub

Private Sub lbl_Click(Index As Integer)
  On Error Resume Next
  frmChat.SControl.Run strPrefix & "_" & strFormName & "_" & lbl(Index).Tag & "_Click"
End Sub

Private Sub lbl_DblClick(Index As Integer)
  On Error Resume Next
  frmChat.SControl.Run strPrefix & "_" & strFormName & "_" & lbl(Index).Tag & "_DblClick"
End Sub

Private Sub lbl_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
  frmChat.SControl.Run strPrefix & "_" & strFormName & "_" & lbl(Index).Tag & "_MouseDown", Button, Shift, X, Y
End Sub

Private Sub lbl_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
  frmChat.SControl.Run strPrefix & "_" & strFormName & "_" & lbl(Index).Tag & "_MouseMove", Button, Shift, X, Y
End Sub

Private Sub lbl_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
  frmChat.SControl.Run strPrefix & "_" & strFormName & "_" & lbl(Index).Tag & "_MouseUp", Button, Shift, X, Y
End Sub

Private Sub tmr_Timer(Index As Integer)
  On Error Resume Next
  frmChat.SControl.Run strPrefix & "_" & strFormName & "_" & tmr(Index).Tag & "_Timer"
End Sub

Private Sub txt_Change(Index As Integer)
  On Error Resume Next
  frmChat.SControl.Run strPrefix & "_" & strFormName & "_" & txt(Index).Tag & "_Change"
End Sub

Private Sub txt_Click(Index As Integer)
  On Error Resume Next
  frmChat.SControl.Run strPrefix & "_" & strFormName & "_" & txt(Index).Tag & "_Click"
End Sub

Private Sub txt_DblClick(Index As Integer)
  On Error Resume Next
  frmChat.SControl.Run strPrefix & "_" & strFormName & "_" & txt(Index).Tag & "_DblClick"
End Sub

Private Sub txt_LostFocus(Index As Integer)
  On Error Resume Next
  frmChat.SControl.Run strPrefix & "_" & strFormName & "_" & txt(Index).Tag & "_LostFocus"
End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
  frmChat.SControl.Run strPrefix & "_" & strFormName & "_" & txt(Index).Tag & "_MouseDown", Button, Shift, X, Y
End Sub

Private Sub txt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
  frmChat.SControl.Run strPrefix & "_" & strFormName & "_" & txt(Index).Tag & "_MouseMove", Button, Shift, X, Y
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
  frmChat.SControl.Run strPrefix & "_" & strFormName & "_" & txt(Index).Tag & "_MouseUp", Button, Shift, X, Y
End Sub

Private Sub txt_GotFocus(Index As Integer)
  On Error Resume Next
  frmChat.SControl.Run strPrefix & "_" & strFormName & "_" & txt(Index).Tag & "_GotFocus"
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
  On Error Resume Next
  frmChat.SControl.Run strPrefix & "_" & strFormName & "_" & txt(Index).Tag & "_KeyPress", KeyAscii
End Sub

Private Sub txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
  On Error Resume Next
  frmChat.SControl.Run strPrefix & "_" & strFormName & "_" & txt(Index).Tag & "_KeyUp", KeyCode, Shift
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  On Error Resume Next
  frmChat.SControl.Run strPrefix & "_" & strFormName & "_" & txt(Index).Tag & "_KeyDown", KeyCode, Shift
End Sub

