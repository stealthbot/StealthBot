VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmScriptUI 
   BackColor       =   &H00000000&
   Caption         =   "Scripting UI"
   ClientHeight    =   3195
   ClientLeft      =   840
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox rtb 
      Height          =   255
      Left            =   1560
      TabIndex        =   9
      Top             =   1200
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmScruptUI.frx":0000
   End
   Begin InetCtlsObjects.Inet ine 
      Index           =   0
      Left            =   2280
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList iml 
      Index           =   0
      Left            =   1560
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lsv 
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   8
      Top             =   360
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.ListBox lst 
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ComboBox cmb 
      Height          =   315
      Index           =   0
      Left            =   2040
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.OptionButton opt 
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chk 
      Height          =   255
      Index           =   0
      Left            =   1320
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic 
      Height          =   255
      Index           =   0
      Left            =   960
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Timer tmr 
      Index           =   0
      Left            =   0
      Top             =   360
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   0
      Left            =   600
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
   Begin VB.Line lin 
      Index           =   0
      Visible         =   0   'False
      X1              =   720
      X2              =   1920
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Shape shp 
      Height          =   255
      Index           =   0
      Left            =   2880
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
      Left            =   360
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
Private oNames As Object
Private bSettingsFilled As Boolean

Public Sub FillPrefixName(ByVal sPrefix As String, ByVal strName As String)
  If (bSettingsFilled) Then Exit Sub
  strPrefix = sPrefix
  strFormName = strName
  bSettingsFilled = True
End Sub

Public Function GetControl(ByVal strName As String) As Object
  If (oControls.Exists(strName)) Then
    Set GetControl = oControls.Item(strName)
  Else
    Set GetControl = Nothing
  End If
End Function

Private Function AddControl(ByVal strName As String, ByVal strCtlName As String, ByRef ctrls As Object, Optional bVisable As Boolean = True)
  If (oControls.Exists(strName)) Then
    AddControl = False
  Else
    Dim Index As Integer
    Index = ctrls.UBound + 1
    Load ctrls(Index)
    If (bVisable) Then ctrls(Index).Visible = True
    oNames.Add strCtlName & "_" & Index, strName
    oControls.Add strName, ctrls(Index)
    Form_Resize
    AddControl = True
  End If
End Function

Public Function AddCommandButton(ByVal strName As String) As Boolean
  AddCommandButton = AddControl(strName, "cmd", cmd)
End Function

Public Function AddLabel(ByVal strName As String) As Boolean
  AddLabel = AddControl(strName, "lbl", lbl)
End Function

Public Function AddTextBox(ByVal strName As String) As Boolean
  AddTextBox = AddControl(strName, "txt", txt)
End Function

Public Function AddTimer(ByVal strName As String) As Boolean
  AddTimer = AddControl(strName, "tmr", tmr, False)
End Function

Public Function AddPictureBox(ByVal strName As String) As Boolean
  AddPictureBox = AddControl(strName, "pic", pic)
End Function

Public Function AddCheckBox(ByVal strName As String) As Boolean
  AddCheckBox = AddControl(strName, "chk", chk)
End Function

Public Function AddOptionBox(ByVal strName As String) As Boolean
  AddOptionBox = AddControl(strName, "opt", opt)
End Function

Public Function AddComboBox(ByVal strName As String) As Boolean
  AddComboBox = AddControl(strName, "cmb", cmb)
End Function

Public Function AddListBox(ByVal strName As String) As Boolean
  AddListBox = AddControl(strName, "lst", lst)
End Function

Public Function AddShape(ByVal strName As String) As Boolean
  AddShape = AddControl(strName, "shp", shp)
End Function

Public Function AddLine(ByVal strName As String) As Boolean
  AddLine = AddControl(strName, "lin", lin)
End Function

Public Function AddListView(ByVal strName As String) As Boolean
  AddListView = AddControl(strName, "lsv", lsv)
End Function

Public Function AddImageList(ByVal strName As String) As Boolean
  AddImageList = AddControl(strName, "iml", iml, False)
End Function

Public Function AddINet(ByVal strName As String) As Boolean
  AddINet = AddControl(strName, "ine", ine, False)
End Function

Public Function AddRichTextBox(ByVal strName As String) As Boolean
  AddINet = AddControl(strName, "rtb", rtb)
End Function

Public Sub DestroyObjects()
  On Error Resume Next
  Dim x As Integer
  For x = 1 To lbl.UBound: Unload lbl(x): Next x
  For x = 1 To cmd.UBound: Unload cmd(x): Next x
  For x = 1 To txt.UBound: Unload txt(x): Next x
  For x = 1 To tmr.UBound: Unload tmr(x): Next x
  For x = 1 To pic.UBound: Unload pic(x): Next x
  For x = 1 To chk.UBound: Unload chk(x): Next x
  For x = 1 To opt.UBound: Unload opt(x): Next x
  For x = 1 To cmb.UBound: Unload cmb(x): Next x
  For x = 1 To lst.UBound: Unload lst(x): Next x
  For x = 1 To shp.UBound: Unload shp(x): Next x
  For x = 1 To lin.UBound: Unload lin(x): Next x
  For x = 1 To lsv.UBound: Unload lsv(x): Next x
  For x = 1 To iml.UBound: Unload iml(x): Next x
  For x = 1 To ine.UBound: Unload ine(x): Next x
  For x = 1 To rtb.UBound: Unload rtb(x): Next x
  Set oControls = Nothing
End Sub

Private Function GetCallBack(ByVal strObject As String, ByVal Index As Integer, ByVal strFunction As String)
  If (strObject = vbNullString) Then
    GetCallBack = strPrefix & "_" & strFormName & "_" & strFunction
    Exit Function
  End If
  If (oNames.Exists(strObject & "_" & Index)) Then
    GetCallBack = strPrefix & "_" & strFormName & "_" & _
    oNames.Item(strObject & "_" & Index) & "_" & strFunction
  Else
    GetCallBack = strPrefix & "_" & strFormName & "_" & _
    strObject & "_" & Index & "_" & strFunction
  End If
End Function

'//////////////////////////////////////////////////////
'//Events
'//////////////////////////////////////////////////////

Private Sub Form_Load()
  Me.Icon = frmChat.Icon
  Set oControls = CreateObject("Scripting.Dictionary")
  Set oNames = CreateObject("Scripting.Dictionary")
  On Error Resume Next
  frmChat.SControl.Run GetCallBack(vbNullString, 0, "Load")
End Sub

Private Sub Form_Activate()
  On Error Resume Next
  frmChat.SControl.Run GetCallBack(vbNullString, 0, "Activate")
End Sub

Private Sub Form_Click()
  On Error Resume Next
  frmChat.SControl.Run GetCallBack(vbNullString, 0, "Click")
End Sub

Private Sub Form_DblClick()
  On Error Resume Next
  frmChat.SControl.Run GetCallBack(vbNullString, 0, "DblClick")
End Sub

Private Sub Form_Deactivate()
  On Error Resume Next
  frmChat.SControl.Run GetCallBack(vbNullString, 0, "Deactivate")
End Sub

Private Sub Form_GotFocus()
  On Error Resume Next
  frmChat.SControl.Run GetCallBack(vbNullString, 0, "GotFocus")
End Sub

Private Sub Form_Initialize()
  On Error Resume Next
  frmChat.SControl.Run GetCallBack(vbNullString, 0, "Initialize")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error Resume Next
  frmChat.SControl.Run GetCallBack(vbNullString, 0, "KeyDown"), KeyCode, Shift
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  On Error Resume Next
  frmChat.SControl.Run GetCallBack(vbNullString, 0, "KeyPress"), KeyAscii
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  On Error Resume Next
  frmChat.SControl.Run GetCallBack(vbNullString, 0, "KeyUp"), KeyCode, Shift
End Sub

Private Sub Form_LostFocus()
  On Error Resume Next
  frmChat.SControl.Run GetCallBack(vbNullString, 0, "LostFocus")
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  On Error Resume Next
  frmChat.SControl.Run GetCallBack(vbNullString, 0, "MouseDown"), Button, Shift, x, Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  On Error Resume Next
  frmChat.SControl.Run GetCallBack(vbNullString, 0, "MoveMouse"), Button, Shift, x, Y
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  On Error Resume Next
  frmChat.SControl.Run GetCallBack(vbNullString, 0, "MouseUp"), Button, Shift, x, Y
End Sub

Private Sub Form_Paint()
  On Error Resume Next
  frmChat.SControl.Run GetCallBack(vbNullString, 0, "Paint")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error Resume Next
  frmChat.SControl.Run GetCallBack(vbNullString, 0, "QueryUnload"), Cancel, UnloadMode
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  frmChat.SControl.Run GetCallBack(vbNullString, 0, "Resize")
End Sub

Private Sub Form_Terminate()
  On Error Resume Next
  frmChat.SControl.Run GetCallBack(vbNullString, 0, "Terminate")
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  frmChat.SControl.Run GetCallBack(vbNullString, 0, "Unload"), Cancel
  frmChat.SControl.ExecuteStatement "Call DestroyForm(" & Chr(&H22) & strPrefix & Chr(&H22) & ", " & Chr(&H22) & strFormName & Chr(&H22) & ")"
End Sub

Private Sub cmd_LostFocus(Index As Integer)
  On Error Resume Next
  frmChat.SControl.Run GetCallBack("cmd", Index, "LostFocus")
End Sub

Private Sub cmd_GotFocus(Index As Integer)
  On Error Resume Next
  frmChat.SControl.Run GetCallBack("cmd", Index, "GotFocus")
End Sub

Private Sub cmd_KeyPress(Index As Integer, KeyAscii As Integer)
  On Error Resume Next
  frmChat.SControl.Run GetCallBack("cmd", Index, "KeyPress"), KeyAscii
End Sub

Private Sub cmd_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
  On Error Resume Next
  frmChat.SControl.Run GetCallBack("cmd", Index, "KeyUp"), KeyCode, Shift
End Sub

Private Sub cmd_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
  On Error Resume Next
  frmChat.SControl.Run GetCallBack("cmd", Index, "MouseDown"), Button, Shift, x, Y
End Sub

Private Sub cmd_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
  On Error Resume Next
  frmChat.SControl.Run GetCallBack("cmd", Index, "MouseMove"), Button, Shift, x, Y
End Sub

Private Sub cmd_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
  On Error Resume Next
  frmChat.SControl.Run GetCallBack("cmd", Index, "MouseUp"), Button, Shift, x, Y
End Sub

Private Sub cmd_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  On Error Resume Next
  frmChat.SControl.Run GetCallBack("cmd", Index, "KeyDown"), KeyCode, Shift
End Sub

Private Sub cmd_Click(Index As Integer)
  On Error Resume Next
  frmChat.SControl.Run GetCallBack("cmd", Index, "Click")
End Sub

Private Sub lbl_Change(Index As Integer)
  On Error Resume Next
  frmChat.SControl.Run GetCallBack("lbl", Index, "Change")
End Sub

Private Sub lbl_Click(Index As Integer)
  On Error Resume Next
  frmChat.SControl.Run GetCallBack("lbl", Index, "Click")
End Sub

Private Sub lbl_DblClick(Index As Integer)
  On Error Resume Next
  frmChat.SControl.Run GetCallBack("lbl", Index, "DblClick")
End Sub

Private Sub lbl_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
  On Error Resume Next
  frmChat.SControl.Run GetCallBack("lbl", Index, "MouseDown"), Button, Shift, x, Y
End Sub

Private Sub lbl_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
  On Error Resume Next
  frmChat.SControl.Run GetCallBack("lbl", Index, "MouseMove"), Button, Shift, x, Y
End Sub

Private Sub lbl_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
  On Error Resume Next
  frmChat.SControl.Run GetCallBack("lbl", Index, "MouseUp"), Button, Shift, x, Y
End Sub

Private Sub tmr_Timer(Index As Integer)
  On Error Resume Next
  frmChat.SControl.Run GetCallBack("tmr", Index, "Timer")
End Sub

Private Sub txt_Change(Index As Integer)
  On Error Resume Next
  frmChat.SControl.Run GetCallBack("txt", Index, "Change")
End Sub

Private Sub txt_Click(Index As Integer)
  On Error Resume Next
  frmChat.SControl.Run GetCallBack("txt", Index, "Click")
End Sub

Private Sub txt_DblClick(Index As Integer)
  On Error Resume Next
  frmChat.SControl.Run GetCallBack("txt", Index, "DblClick")
End Sub

Private Sub txt_LostFocus(Index As Integer)
  On Error Resume Next
  frmChat.SControl.Run GetCallBack("txt", Index, "LostFocus")
End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
  On Error Resume Next
  frmChat.SControl.Run GetCallBack("txt", Index, "MouseDown"), Button, Shift, x, Y
End Sub

Private Sub txt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
  On Error Resume Next
  frmChat.SControl.Run GetCallBack("txt", Index, "MouseMove"), Button, Shift, x, Y
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
  On Error Resume Next
  frmChat.SControl.Run GetCallBack("txt", Index, "MouseUp"), Button, Shift, x, Y
End Sub

Private Sub txt_GotFocus(Index As Integer)
  On Error Resume Next
  frmChat.SControl.Run GetCallBack("txt", Index, "GotFocus")
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
  On Error Resume Next
  frmChat.SControl.Run GetCallBack("txt", Index, "KeyPress"), KeyAscii
End Sub

Private Sub txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
  On Error Resume Next
  frmChat.SControl.Run GetCallBack("txt", Index, "KeyUp"), KeyCode, Shift
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  On Error Resume Next
  frmChat.SControl.Run GetCallBack("txt", Index, "KeyDown"), KeyCode, Shift
End Sub

