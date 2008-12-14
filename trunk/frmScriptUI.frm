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
    Dim Index As Integer
    
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
        Index = (ctrls.UBound + 1)
        
        ' ...
        Call Load(ctrls(Index))
        
        ' ...
        If (bVisable) Then
            ctrls(Index).Visible = True
        End If
        
        ' ...
        Call oNames.Add(strCtlName & "_" & Index, strName)
        
        ' ...
        Call oControls.Add(strName, ctrls(Index))
        
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

Private Function GetCallBack(ByVal strObject As String, ByVal Index As Integer, ByVal strFunction As String)
    If (strObject = vbNullString) Then
        GetCallBack = strPrefix & "_" & strFormName & "_" & strFunction
        
        Exit Function
    End If
    
    If (oNames.Exists(strObject & "_" & Index)) Then
        GetCallBack = strPrefix & "_" & strFormName & "_" & oNames.Item(strObject & _
            "_" & Index) & "_" & strFunction
    Else
        GetCallBack = strPrefix & "_" & strFormName & "_" & strObject & "_" & Index & _
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

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack(vbNullString, 0, "MouseDown"), Button, Shift, X, Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack(vbNullString, 0, "MoveMouse"), Button, Shift, X, Y
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack(vbNullString, 0, "MouseUp"), Button, Shift, X, Y
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

Private Sub cmd_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("cmd", Index, "MouseDown"), Button, Shift, X, Y
End Sub

Private Sub cmd_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("cmd", Index, "MouseMove"), Button, Shift, X, Y
End Sub

Private Sub cmd_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("cmd", Index, "MouseUp"), Button, Shift, X, Y
End Sub

Private Sub cmd_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("cmd", Index, "KeyDown"), KeyCode, Shift
End Sub

Private Sub cmd_Click(Index As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("cmd", Index, "Click")
End Sub

Private Sub ine_StateChanged(Index As Integer, ByVal State As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("ine", Index, "StateChanged"), State
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

Private Sub lbl_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("lbl", Index, "MouseDown"), Button, Shift, X, Y
End Sub

Private Sub lbl_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("lbl", Index, "MouseMove"), Button, Shift, X, Y
End Sub

Private Sub lbl_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("lbl", Index, "MouseUp"), Button, Shift, X, Y
End Sub

Private Sub lst_Click(Index As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("lst", Index, "Click")
End Sub

Private Sub lst_DblClick(Index As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("lst", Index, "DblClick")
End Sub

Private Sub lst_GotFocus(Index As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("lst", Index, "GotFocus")
End Sub

Private Sub lst_ItemCheck(Index As Integer, Item As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("lst", Index, "ItemClick"), Item
End Sub

Private Sub lst_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("lst", Index, "KeyDown"), KeyCode, Shift
End Sub

Private Sub lst_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("lst", Index, "KeyPress"), KeyAscii
End Sub

Private Sub lst_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("lst", Index, "KeyUp"), KeyCode, Shift
End Sub

Private Sub lst_LostFocus(Index As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("lst", Index, "LostFocus")
End Sub

Private Sub lst_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("lst", Index, "MouseDown"), Button, Shift, X, Y
End Sub

Private Sub lst_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("lst", Index, "MouseMove"), Button, Shift, X, Y
End Sub

Private Sub lst_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("lst", Index, "MouseUp"), Button, Shift, X, Y
End Sub

Private Sub lst_Scroll(Index As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("lst", Index, "Scroll")
End Sub

Private Sub lsv_AfterLabelEdit(Index As Integer, Cancel As Integer, NewString As String)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("lsv", Index, "AfterLabelEdit"), Cancel, NewString
End Sub

Private Sub lsv_BeforeLabelEdit(Index As Integer, Cancel As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("lsv", Index, "BeforeLabelEdit"), Cancel
End Sub

Private Sub lsv_Click(Index As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("lsv", Index, "Click")
End Sub

Private Sub lsv_DblClick(Index As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("lsv", Index, "DblClick")
End Sub

Private Sub lsv_GotFocus(Index As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("lsv", Index, "GotFocus")
End Sub

Private Sub lsv_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("lsv", Index, "KeyDown"), KeyCode, Shift
End Sub

Private Sub lsv_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("lsv", Index, "KeyPress")
End Sub

Private Sub lsv_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("lsv", Index, "KeyUp"), KeyCode, Shift
End Sub

Private Sub lsv_LostFocus(Index As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("lsv", Index, "LostFocus")
End Sub

Private Sub lsv_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("lsv", Index, "MouseDown"), Button, Shift, X, Y
End Sub

Private Sub lsv_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("lsv", Index, "MouseMove"), Button, Shift, X, Y
End Sub

Private Sub lsv_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("lsv", Index, "MouseUp"), Button, Shift, X, Y
End Sub

Private Sub opt_Click(Index As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("opt", Index, "Click")
End Sub

Private Sub opt_DblClick(Index As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("opt", Index, "DblClick")
End Sub

Private Sub opt_GotFocus(Index As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("opt", Index, "GotFocus")
End Sub

Private Sub opt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("opt", Index, "KeyDown"), KeyCode, Shift
End Sub

Private Sub opt_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("opt", Index, "KeyPress"), KeyAscii
End Sub

Private Sub opt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("opt", Index, "KeyUp"), KeyCode, Shift
End Sub

Private Sub opt_LostFocus(Index As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("opt", Index, "LostFocus")
End Sub

Private Sub opt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("opt", Index, "MouseDown"), Button, Shift, X, Y
End Sub

Private Sub opt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("opt", Index, "MouseMove"), Button, Shift, X, Y
End Sub

Private Sub opt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("opt", Index, "MouseUp"), Button, Shift, X, Y
End Sub

Private Sub pic_Change(Index As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("pic", Index, "Change")
End Sub

Private Sub pic_Click(Index As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("pic", Index, "Click")
End Sub

Private Sub pic_DblClick(Index As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("pic", Index, "DblClick")
End Sub

Private Sub pic_GotFocus(Index As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("pic", Index, "GotFocus")
End Sub

Private Sub pic_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("pic", Index, "KeyDown"), KeyCode, Shift
End Sub

Private Sub pic_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("pic", Index, "KeyPress"), KeyAscii
End Sub

Private Sub pic_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("pic", Index, "KeyUp"), KeyCode, Shift
End Sub

Private Sub pic_LinkClose(Index As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("pic", Index, "LinkClose")
End Sub

Private Sub pic_LinkError(Index As Integer, LinkErr As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("pic", Index, "LinkError"), LinkErr
End Sub

Private Sub pic_LinkNotify(Index As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("pic", Index, "LinkNotify")
End Sub

Private Sub pic_LinkOpen(Index As Integer, Cancel As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("pic", Index, "LinkOpen"), Cancel
End Sub

Private Sub pic_LostFocus(Index As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("pic", Index, "LostFocus")
End Sub

Private Sub pic_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("pic", Index, "MouseDown"), Button, Shift, X, Y
End Sub

Private Sub pic_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("pic", Index, "MouseMove"), Button, Shift, X, Y
End Sub

Private Sub pic_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("pic", Index, "MouseUp"), Button, Shift, X, Y
End Sub

Private Sub pic_Paint(Index As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("pic", Index, "Paint")
End Sub

Private Sub pic_Resize(Index As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("pic", Index, "Resize")
End Sub

Private Sub rtb_Change(Index As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("rtb", Index, "Change")
End Sub

Private Sub rtb_Click(Index As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("rtb", Index, "Click")
End Sub

Private Sub rtb_DblClick(Index As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("rtb", Index, "DblClick")
End Sub

Private Sub rtb_GotFocus(Index As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("rtb", Index, "GotFocus")
End Sub

Private Sub rtb_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("rtb", Index, "KeyDown"), KeyCode, Shift
End Sub

Private Sub rtb_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("rtb", Index, "KeyPress"), KeyAscii
End Sub

Private Sub rtb_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("rtb", Index, "KeyUp"), KeyCode, Shift
End Sub

Private Sub rtb_LostFocus(Index As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("rtb", Index, "LostFocus")
End Sub

Private Sub rtb_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("rtb", Index, "MouseDown"), Button, Shift, X, Y
End Sub

Private Sub rtb_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("rtb", Index, "MouseMove"), Button, Shift, X, Y
End Sub

Private Sub rtb_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("rtb", Index, "MouseUp"), Button, Shift, X, Y
End Sub

Private Sub rtb_SelChange(Index As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("rtb", Index, "SelChange")
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

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("txt", Index, "MouseDown"), Button, Shift, X, Y
End Sub

Private Sub txt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("txt", Index, "MouseMove"), Button, Shift, X, Y
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("txt", Index, "MouseUp"), Button, Shift, X, Y
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

Private Sub chk_Click(Index As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("chk", Index, "Click")
End Sub

Private Sub chk_GotFocus(Index As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("chk", Index, "GotFocus")
End Sub

Private Sub chk_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("chk", Index, "KeyDown"), KeyCode, Shift
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("chk", Index, "KeyPress"), KeyAscii
End Sub

Private Sub chk_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("chk", Index, "KeyUp"), KeyCode, Shift
End Sub

Private Sub chk_LostFocus(Index As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("chk", Index, "LostFocus")
End Sub

Private Sub chk_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("chk", Index, "MouseDown"), Button, Shift, X, Y
End Sub

Private Sub chk_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("chk", Index, "MouseMove"), Button, Shift, X, Y
End Sub

Private Sub chk_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("chk", Index, "MouseUp"), Button, Shift, X, Y
End Sub

Private Sub cmb_Change(Index As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("cmb", Index, "Change")
End Sub

Private Sub cmb_Click(Index As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("cmb", Index, "Click")
End Sub

Private Sub cmb_DblClick(Index As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("cmb", Index, "DblClick")
End Sub

Private Sub cmb_GotFocus(Index As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("cmb", Index, "GotFocus")
End Sub

Private Sub cmb_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("cmb", Index, "KeyDown"), KeyCode, Shift
End Sub

Private Sub cmb_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("cmb", Index, "KeyPress"), KeyAscii
End Sub

Private Sub cmb_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("cmb", Index, "KeyUp"), KeyCode, Shift
End Sub

Private Sub cmb_LostFocus(Index As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("cmb", Index, "LostFocus")
End Sub

Private Sub cmb_Scroll(Index As Integer)
    On Error Resume Next
    
    frmChat.SControl.Run GetCallBack("cmb", Index, "Scroll")
End Sub

Public Sub AddChat(ParamArray saElements() As Variant)
    Dim arr() As Variant ' ...
    
    ' ...
    arr() = saElements
    
    ' ...
    Call DisplayRichText(frmScriptUI.rtb, arr)
End Sub
