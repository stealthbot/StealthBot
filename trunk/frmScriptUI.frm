VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmScriptUI 
   BackColor       =   &H00000000&
   Caption         =   "Scripting UI"
   ClientHeight    =   3135
   ClientLeft      =   975
   ClientTop       =   360
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4590
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox rtb 
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmScriptUI.frx":0000
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
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      OLEDragMode     =   1
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDragMode     =   1
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
      MultiLine       =   -1  'True
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
' frmScriptUI.frm
' Copyright 2007 Hdx
' Create and manage form from scripting system

Option Explicit

Private strPrefix       As String
Private strFormName     As String
Private oControls       As Object
Private oNames          As Object
Private bSettingsFilled As Boolean

Public Sub FillPrefixName(ByVal sPrefix As String, ByVal strName As String)
    If (bSettingsFilled) Then
        Exit Sub
    End If
    
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
    Dim index As Integer
    
    ' ...
    If (oControls Is Nothing) Then
        Set oControls = CreateObject("Scripting.Dictionary")
    End If
    
    ' ...
    If (oNames Is Nothing) Then
        Set oNames = CreateObject("Scripting.Dictionary")
    End If
    
    ' does control already exist?
    If (oControls.Exists(strName)) Then
        ' alert calling procedure of failure
        AddControl = False
    
        ' break from function
        Exit Function
    Else
        ' ...
        index = (ctrls.UBound + 1)
        
        ' ...
        Call Load(ctrls(index))
        
        ' ...
        If (bVisable) Then
            ctrls(index).Visible = True
        End If
        
        ' ...
        Call oNames.Add(strCtlName & "_" & index, strName)
        
        ' ...
        Call oControls.Add(strName, ctrls(index))
        
        ' ...
        Call Form_Resize
        
        ' ...
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
    AddRichTextBox = AddControl(strName, "rtb", rtb)
End Function

Public Sub DestroyObjects()
    On Error Resume Next
    
    Dim X As Integer
    
    For X = 1 To lbl.UBound: Unload lbl(X): Next X
    For X = 1 To cmd.UBound: Unload cmd(X): Next X
    For X = 1 To txt.UBound: Unload txt(X): Next X
    For X = 1 To tmr.UBound: Unload tmr(X): Next X
    For X = 1 To pic.UBound: Unload pic(X): Next X
    For X = 1 To chk.UBound: Unload chk(X): Next X
    For X = 1 To opt.UBound: Unload opt(X): Next X
    For X = 1 To cmb.UBound: Unload cmb(X): Next X
    For X = 1 To lst.UBound: Unload lst(X): Next X
    For X = 1 To shp.UBound: Unload shp(X): Next X
    For X = 1 To lin.UBound: Unload lin(X): Next X
    For X = 1 To lsv.UBound: Unload lsv(X): Next X
    For X = 1 To iml.UBound: Unload iml(X): Next X
    For X = 1 To ine.UBound: Unload ine(X): Next X
    For X = 1 To rtb.UBound: Unload rtb(X): Next X
    
    Set oControls = Nothing
End Sub

Private Function GetCallBack(ByVal strObject As String, ByVal index As Integer, ByVal strFunction As String)
    If (strObject = vbNullString) Then
        GetCallBack = strPrefix & "_" & strFormName & "_" & strFunction
        
        Exit Function
    End If
    
    If (oNames.Exists(strObject & "_" & index)) Then
        GetCallBack = strPrefix & "_" & strFormName & "_" & oNames.Item(strObject & _
            "_" & index) & "_" & strFunction
    Else
        GetCallBack = strPrefix & "_" & strFormName & "_" & strObject & "_" & index & _
            "_" & strFunction
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
    
    RunInAll frmChat.SControl, GetCallBack(vbNullString, 0, "Load")
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack(vbNullString, 0, "Activate")
End Sub

Private Sub Form_Click()
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack(vbNullString, 0, "Click")
End Sub

Private Sub Form_DblClick()
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack(vbNullString, 0, "DblClick")
End Sub

Private Sub Form_Deactivate()
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack(vbNullString, 0, "Deactivate")
End Sub

Private Sub Form_GotFocus()
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack(vbNullString, 0, "GotFocus")
End Sub

Private Sub Form_Initialize()
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack(vbNullString, 0, "Initialize")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack(vbNullString, 0, "KeyDown"), KeyCode, Shift
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack(vbNullString, 0, "KeyPress"), KeyAscii
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack(vbNullString, 0, "KeyUp"), KeyCode, Shift
End Sub

Private Sub Form_LostFocus()
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack(vbNullString, 0, "LostFocus")
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack(vbNullString, 0, "MouseDown"), Button, Shift, X, Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack(vbNullString, 0, "MoveMouse"), Button, Shift, X, Y
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack(vbNullString, 0, "MouseUp"), Button, Shift, X, Y
End Sub

Private Sub Form_Paint()
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack(vbNullString, 0, "Paint")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack(vbNullString, 0, "QueryUnload"), Cancel, UnloadMode
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack(vbNullString, 0, "Resize")
End Sub

Private Sub Form_Terminate()
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack(vbNullString, 0, "Terminate")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack(vbNullString, 0, "Unload"), Cancel
    frmChat.SControl.ExecuteStatement "Call DestroyForm(" & Chr(&H22) & strPrefix & Chr(&H22) & ", " & Chr(&H22) & strFormName & Chr(&H22) & ")"
End Sub

Private Sub cmd_LostFocus(index As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("cmd", index, "LostFocus")
End Sub

Private Sub cmd_GotFocus(index As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("cmd", index, "GotFocus")
End Sub

Private Sub cmd_KeyPress(index As Integer, KeyAscii As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("cmd", index, "KeyPress"), KeyAscii
End Sub

Private Sub cmd_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("cmd", index, "KeyUp"), KeyCode, Shift
End Sub

Private Sub cmd_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("cmd", index, "MouseDown"), Button, Shift, X, Y
End Sub

Private Sub cmd_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("cmd", index, "MouseMove"), Button, Shift, X, Y
End Sub

Private Sub cmd_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("cmd", index, "MouseUp"), Button, Shift, X, Y
End Sub

Private Sub cmd_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("cmd", index, "KeyDown"), KeyCode, Shift
End Sub

Private Sub cmd_Click(index As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("cmd", index, "Click")
End Sub

Private Sub ine_StateChanged(index As Integer, ByVal State As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("ine", index, "StateChanged"), State
End Sub

Private Sub lbl_Change(index As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("lbl", index, "Change")
End Sub

Private Sub lbl_Click(index As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("lbl", index, "Click")
End Sub

Private Sub lbl_DblClick(index As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("lbl", index, "DblClick")
End Sub

Private Sub lbl_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("lbl", index, "MouseDown"), Button, Shift, X, Y
End Sub

Private Sub lbl_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("lbl", index, "MouseMove"), Button, Shift, X, Y
End Sub

Private Sub lbl_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("lbl", index, "MouseUp"), Button, Shift, X, Y
End Sub

Private Sub lst_Click(index As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("lst", index, "Click")
End Sub

Private Sub lst_DblClick(index As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("lst", index, "DblClick")
End Sub

Private Sub lst_GotFocus(index As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("lst", index, "GotFocus")
End Sub

Private Sub lst_ItemCheck(index As Integer, Item As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("lst", index, "ItemClick"), Item
End Sub

Private Sub lst_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("lst", index, "KeyDown"), KeyCode, Shift
End Sub

Private Sub lst_KeyPress(index As Integer, KeyAscii As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("lst", index, "KeyPress"), KeyAscii
End Sub

Private Sub lst_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("lst", index, "KeyUp"), KeyCode, Shift
End Sub

Private Sub lst_LostFocus(index As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("lst", index, "LostFocus")
End Sub

Private Sub lst_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("lst", index, "MouseDown"), Button, Shift, X, Y
End Sub

Private Sub lst_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("lst", index, "MouseMove"), Button, Shift, X, Y
End Sub

Private Sub lst_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("lst", index, "MouseUp"), Button, Shift, X, Y
End Sub

Private Sub lst_Scroll(index As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("lst", index, "Scroll")
End Sub

Private Sub lsv_AfterLabelEdit(index As Integer, Cancel As Integer, NewString As String)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("lsv", index, "AfterLabelEdit"), Cancel, NewString
End Sub

Private Sub lsv_BeforeLabelEdit(index As Integer, Cancel As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("lsv", index, "BeforeLabelEdit"), Cancel
End Sub

Private Sub lsv_Click(index As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("lsv", index, "Click")
End Sub

Private Sub lsv_DblClick(index As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("lsv", index, "DblClick")
End Sub

Private Sub lsv_GotFocus(index As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("lsv", index, "GotFocus")
End Sub

Private Sub lsv_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("lsv", index, "KeyDown"), KeyCode, Shift
End Sub

Private Sub lsv_KeyPress(index As Integer, KeyAscii As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("lsv", index, "KeyPress")
End Sub

Private Sub lsv_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("lsv", index, "KeyUp"), KeyCode, Shift
End Sub

Private Sub lsv_LostFocus(index As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("lsv", index, "LostFocus")
End Sub

Private Sub lsv_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("lsv", index, "MouseDown"), Button, Shift, X, Y
End Sub

Private Sub lsv_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("lsv", index, "MouseMove"), Button, Shift, X, Y
End Sub

Private Sub lsv_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("lsv", index, "MouseUp"), Button, Shift, X, Y
End Sub

Private Sub opt_Click(index As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("opt", index, "Click")
End Sub

Private Sub opt_DblClick(index As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("opt", index, "DblClick")
End Sub

Private Sub opt_GotFocus(index As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("opt", index, "GotFocus")
End Sub

Private Sub opt_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("opt", index, "KeyDown"), KeyCode, Shift
End Sub

Private Sub opt_KeyPress(index As Integer, KeyAscii As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("opt", index, "KeyPress"), KeyAscii
End Sub

Private Sub opt_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("opt", index, "KeyUp"), KeyCode, Shift
End Sub

Private Sub opt_LostFocus(index As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("opt", index, "LostFocus")
End Sub

Private Sub opt_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("opt", index, "MouseDown"), Button, Shift, X, Y
End Sub

Private Sub opt_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("opt", index, "MouseMove"), Button, Shift, X, Y
End Sub

Private Sub opt_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("opt", index, "MouseUp"), Button, Shift, X, Y
End Sub

Private Sub pic_Change(index As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("pic", index, "Change")
End Sub

Private Sub pic_Click(index As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("pic", index, "Click")
End Sub

Private Sub pic_DblClick(index As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("pic", index, "DblClick")
End Sub

Private Sub pic_GotFocus(index As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("pic", index, "GotFocus")
End Sub

Private Sub pic_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("pic", index, "KeyDown"), KeyCode, Shift
End Sub

Private Sub pic_KeyPress(index As Integer, KeyAscii As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("pic", index, "KeyPress"), KeyAscii
End Sub

Private Sub pic_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("pic", index, "KeyUp"), KeyCode, Shift
End Sub

Private Sub pic_LinkClose(index As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("pic", index, "LinkClose")
End Sub

Private Sub pic_LinkError(index As Integer, LinkErr As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("pic", index, "LinkError"), LinkErr
End Sub

Private Sub pic_LinkNotify(index As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("pic", index, "LinkNotify")
End Sub

Private Sub pic_LinkOpen(index As Integer, Cancel As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("pic", index, "LinkOpen"), Cancel
End Sub

Private Sub pic_LostFocus(index As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("pic", index, "LostFocus")
End Sub

Private Sub pic_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("pic", index, "MouseDown"), Button, Shift, X, Y
End Sub

Private Sub pic_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("pic", index, "MouseMove"), Button, Shift, X, Y
End Sub

Private Sub pic_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("pic", index, "MouseUp"), Button, Shift, X, Y
End Sub

Private Sub pic_Paint(index As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("pic", index, "Paint")
End Sub

Private Sub pic_Resize(index As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("pic", index, "Resize")
End Sub

Private Sub rtb_Change(index As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("rtb", index, "Change")
End Sub

Private Sub rtb_Click(index As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("rtb", index, "Click")
End Sub

Private Sub rtb_DblClick(index As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("rtb", index, "DblClick")
End Sub

Private Sub rtb_GotFocus(index As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("rtb", index, "GotFocus")
End Sub

Private Sub rtb_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("rtb", index, "KeyDown"), KeyCode, Shift
End Sub

Private Sub rtb_KeyPress(index As Integer, KeyAscii As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("rtb", index, "KeyPress"), KeyAscii
End Sub

Private Sub rtb_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("rtb", index, "KeyUp"), KeyCode, Shift
End Sub

Private Sub rtb_LostFocus(index As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("rtb", index, "LostFocus")
End Sub

Private Sub rtb_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("rtb", index, "MouseDown"), Button, Shift, X, Y
End Sub

Private Sub rtb_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("rtb", index, "MouseMove"), Button, Shift, X, Y
End Sub

Private Sub rtb_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("rtb", index, "MouseUp"), Button, Shift, X, Y
End Sub

Private Sub rtb_SelChange(index As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("rtb", index, "SelChange")
End Sub

Private Sub tmr_Timer(index As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("tmr", index, "Timer")
End Sub

Private Sub txt_Change(index As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("txt", index, "Change")
End Sub

Private Sub txt_Click(index As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("txt", index, "Click")
End Sub

Private Sub txt_DblClick(index As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("txt", index, "DblClick")
End Sub

Private Sub txt_LostFocus(index As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("txt", index, "LostFocus")
End Sub

Private Sub txt_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("txt", index, "MouseDown"), Button, Shift, X, Y
End Sub

Private Sub txt_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("txt", index, "MouseMove"), Button, Shift, X, Y
End Sub

Private Sub txt_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("txt", index, "MouseUp"), Button, Shift, X, Y
End Sub

Private Sub txt_GotFocus(index As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("txt", index, "GotFocus")
End Sub

Private Sub txt_KeyPress(index As Integer, KeyAscii As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("txt", index, "KeyPress"), KeyAscii
End Sub

Private Sub txt_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("txt", index, "KeyUp"), KeyCode, Shift
End Sub

Private Sub txt_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("txt", index, "KeyDown"), KeyCode, Shift
End Sub

Private Sub chk_Click(index As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("chk", index, "Click")
End Sub

Private Sub chk_GotFocus(index As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("chk", index, "GotFocus")
End Sub

Private Sub chk_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("chk", index, "KeyDown"), KeyCode, Shift
End Sub

Private Sub chk_KeyPress(index As Integer, KeyAscii As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("chk", index, "KeyPress"), KeyAscii
End Sub

Private Sub chk_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("chk", index, "KeyUp"), KeyCode, Shift
End Sub

Private Sub chk_LostFocus(index As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("chk", index, "LostFocus")
End Sub

Private Sub chk_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("chk", index, "MouseDown"), Button, Shift, X, Y
End Sub

Private Sub chk_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("chk", index, "MouseMove"), Button, Shift, X, Y
End Sub

Private Sub chk_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("chk", index, "MouseUp"), Button, Shift, X, Y
End Sub

Private Sub cmb_Change(index As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("cmb", index, "Change")
End Sub

Private Sub cmb_Click(index As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("cmb", index, "Click")
End Sub

Private Sub cmb_DblClick(index As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("cmb", index, "DblClick")
End Sub

Private Sub cmb_GotFocus(index As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("cmb", index, "GotFocus")
End Sub

Private Sub cmb_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("cmb", index, "KeyDown"), KeyCode, Shift
End Sub

Private Sub cmb_KeyPress(index As Integer, KeyAscii As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("cmb", index, "KeyPress"), KeyAscii
End Sub

Private Sub cmb_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("cmb", index, "KeyUp"), KeyCode, Shift
End Sub

Private Sub cmb_LostFocus(index As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("cmb", index, "LostFocus")
End Sub

Private Sub cmb_Scroll(index As Integer)
    On Error Resume Next
    
    RunInAll frmChat.SControl, GetCallBack("cmb", index, "Scroll")
End Sub

Public Sub AddChat(ParamArray saElements() As Variant)
    Dim arr() As Variant ' ...
    
    ' ...
    arr() = saElements
    
    ' ...
    Call DisplayRichText(frmScriptUI.rtb, arr)
End Sub
