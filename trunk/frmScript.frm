VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmScript 
   BackColor       =   &H00000000&
   Caption         =   "Scripting UI"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
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
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar prg 
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   11
      Top             =   360
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.TreeView trv 
      Height          =   615
      Index           =   0
      Left            =   2880
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1085
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin VB.CommandButton cmd 
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   0
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   7
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
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chk 
      Height          =   255
      Index           =   0
      Left            =   1320
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.OptionButton opt 
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ComboBox cmb 
      Height          =   315
      Index           =   0
      Left            =   2040
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.ListBox lst 
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin RichTextLib.RichTextBox rtb 
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmScript.frx":0000
   End
   Begin MSComctlLib.ImageList iml 
      Index           =   0
      Left            =   960
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
      Left            =   0
      TabIndex        =   1
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
   Begin VB.Frame fra 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   2040
      TabIndex        =   12
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lbl 
      BackColor       =   &H00000000&
      Caption         =   "lbl"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape shp 
      Height          =   255
      Index           =   0
      Left            =   2880
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
   Begin VB.Menu dummy 
      Caption         =   "dummy"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_name      As String
Private m_sc_module As Module
Private m_arrObjs() As modScripting.scObj
Private m_objCount  As Integer
Private m_hidden    As Boolean

Public Function SetName(ByVal str As String)

    If (m_name = vbNullString) Then
        m_name = str
    End If

End Function

Public Function GetName() As String

    GetName = m_name

End Function

Public Function SetSCModule(ByRef SCModule As Module)

    If (m_sc_module Is Nothing) Then
        Set m_sc_module = SCModule
    End If

End Function

Public Function GetScriptModule() As Module

    Set GetScriptModule = m_sc_module

End Function

'// 6/22/2009 JSM - Adding wrapper function for MsgBox inside VB6 rather than
'//                 the scripting control. This keeps the focus on the form.
' made parameters optional, like the VBs equivalents -Ribose/2009-08-10
Public Function ShowMsgBox(ByVal Text As String, Optional ByVal opts As VbMsgBoxStyle = vbOKOnly, _
        Optional ByVal Title As String = vbNullString) As VbMsgBoxResult

    ShowMsgBox = MsgBox(Text, opts, Title)

End Function

' wrapper function for InputBox, too!
' made parameters optional, like the VBs equivalents -Ribose/2009-08-10
Public Function ShowInputBox(ByVal Text As String, Optional ByVal Title As String = vbNullString, _
        Optional ByVal Default As String = vbNullString, Optional ByVal XPos As Integer = -1, _
        Optional ByVal YPos As Integer = -1) As String

    If XPos = -1 And YPos = -1 Then
        ShowInputBox = InputBox(Text, Title, Default)
    ElseIf XPos = -1 Then
        ShowInputBox = InputBox(Text, Title, Default, , YPos)
    ElseIf YPos = -1 Then
        ShowInputBox = InputBox(Text, Title, Default, XPos)
    Else
        ShowInputBox = InputBox(Text, Title, Default, XPos, YPos)
    End If
    
End Function

Public Sub DrawLine(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, Optional ByVal Color As Long = -1, Optional ByVal DrawRect As Boolean = False, Optional ByVal FillRect As Boolean = False)

    If Color = -1 Then
        If DrawRect Then
            If FillRect Then
                Line (x1, y1)-(x2, y2), , BF
            Else
                Line (x1, y1)-(x2, y2), , B
            End If
        Else
            Line (x1, y1)-(x2, y2)
        End If
    Else
        If DrawRect Then
            If FillRect Then
                Line (x1, y1)-(x2, y2), Color, BF
            Else
                Line (x1, y1)-(x2, y2), Color, B
            End If
        Else
            Line (x1, y1)-(x2, y2), Color
        End If
    End If

End Sub

'Public Function Objects(objIndex As Integer) As scObj
'
'    Objects = m_arrObjs(objIndex)
'
'End Function

Public Function ObjCount(Optional ObjType As String) As Integer
    Dim i As Integer
    If (ObjType <> vbNullString) Then
        For i = 0 To m_objCount - 1
            If (StrComp(ObjType, m_arrObjs(i).ObjType, vbTextCompare) = 0) Then
                ObjCount = (ObjCount + 1)
            End If
        Next i
    Else
        ObjCount = m_objCount
    End If
End Function

Public Function CreateObj(ByVal ObjType As String, ByVal ObjName As String) As Object
    On Error Resume Next

    Dim obj As scObj
    
    Set CreateObj = Nothing
    If (Not ValidObjectName(ObjName)) Then Exit Function
    
    ' redefine array size & check for duplicate controls
    If (m_objCount) Then
        Dim i As Integer ' loop counter variable

        For i = 0 To m_objCount - 1
            If (StrComp(m_arrObjs(i).ObjType, ObjType, vbTextCompare) = 0) Then
                If (StrComp(m_arrObjs(i).ObjName, ObjName, vbTextCompare) = 0) Then
                    Set CreateObj = m_arrObjs(i).obj
                
                    Exit Function
                End If
            End If
        Next i
        
        ReDim Preserve m_arrObjs(0 To m_objCount)
    Else
        ReDim m_arrObjs(0)
    End If

    Select Case (UCase$(ObjType))
        Case "BUTTON"
            If (ObjCount(ObjType) > 0) Then
                Load cmd(ObjCount(ObjType))
            End If
            
            Set obj.obj = cmd(ObjCount(ObjType))
        
        Case "CHECKBOX"
            If (ObjCount(ObjType) > 0) Then
                Load chk(ObjCount(ObjType))
            End If
            
            Set obj.obj = chk(ObjCount(ObjType))
        
        Case "COMBOBOX"
            If (ObjCount(ObjType) > 0) Then
                Load cmb(ObjCount(ObjType))
            End If
            
            Set obj.obj = cmb(ObjCount(ObjType))
        
        Case "FRAME"
            If (ObjCount(ObjType) > 0) Then
                Load fra(ObjCount(ObjType))
            End If
            
            Set obj.obj = fra(ObjCount(ObjType))
        
        Case "IMAGELIST"
            If (ObjCount(ObjType) > 0) Then
                Load iml(ObjCount(ObjType))
            End If
            
            Set obj.obj = iml(ObjCount(ObjType))
        
        Case "LABEL"
            If (ObjCount(ObjType) > 0) Then
                Load lbl(ObjCount(ObjType))
            End If
            
            Set obj.obj = lbl(ObjCount(ObjType))
        
        Case "LISTBOX"
            If (ObjCount(ObjType) > 0) Then
                Load lst(ObjCount(ObjType))
            End If
            
            Set obj.obj = lst(ObjCount(ObjType))
        
        Case "LISTVIEW"
            If (ObjCount(ObjType) > 0) Then
                Load lsv(ObjCount(ObjType))
            End If
            
            Set obj.obj = lsv(ObjCount(ObjType))
            
        Case "MENU"
            Set obj.obj = New clsMenuObj
            
            obj.obj.Name = GetName() & "_" & ObjName
            
            obj.obj.Parent = Me
            
            DynamicMenus.Add obj.obj
        
        Case "OPTIONBUTTON"
            If (ObjCount(ObjType) > 0) Then
                Load opt(ObjCount(ObjType))
            End If
            
            Set obj.obj = opt(ObjCount(ObjType))
        
        Case "PICTUREBOX"
            If (ObjCount(ObjType) > 0) Then
                Load pic(ObjCount(ObjType))
            End If
            
            Set obj.obj = pic(ObjCount(ObjType))
        
        Case "PROGRESSBAR"
            If (ObjCount(ObjType) > 0) Then
                Load prg(ObjCount(ObjType))
            End If
            
            Set obj.obj = prg(ObjCount(ObjType))
        
        Case "RICHTEXTBOX"
            If (ObjCount(ObjType) > 0) Then
                Load rtb(ObjCount(ObjType))
            End If
            
            Set obj.obj = rtb(ObjCount(ObjType))
            
            EnableURLDetect obj.obj.hWnd
        
        Case "TEXTBOX"
            If (ObjCount(ObjType) > 0) Then
                Load txt(ObjCount(ObjType))
            End If
            
            Set obj.obj = txt(ObjCount(ObjType))
            
        Case "TREEVIEW"
            If (ObjCount(ObjType) > 0) Then
                Load trv(ObjCount(ObjType))
            End If
            
            Set obj.obj = trv(ObjCount(ObjType))
            
    End Select
    
    obj.obj.Visible = True

    ' store our module name & type
    obj.ObjName = ObjName
    obj.ObjType = ObjType

    ' store object
    m_arrObjs(m_objCount) = obj
    
    ' increment object counter
    m_objCount = (m_objCount + 1)

    ' return object
    Set CreateObj = obj.obj
End Function

Public Sub DestroyObjs()

    On Error GoTo ERROR_HANDLER

    Dim i As Integer
    
    For i = m_objCount - 1 To 0 Step -1
        DestroyObj m_arrObjs(i).ObjName
    Next i
    
    Exit Sub

ERROR_HANDLER:
    
    frmChat.AddChat RTBColors.ErrorMessageText, _
        "Error (#" & Err.Number & "): " & Err.Description & " in frmScript::DestroyObjs()."
        
    Resume Next
    
End Sub

Public Sub DestroyObj(ByVal ObjName As String)

    On Error GoTo ERROR_HANDLER

    Dim i     As Integer
    Dim Index As Integer
    
    If (m_objCount = 0) Then
        Exit Sub
    End If
    
    Index = m_objCount
    
    For i = 0 To m_objCount - 1
        If (StrComp(m_arrObjs(i).ObjName, ObjName, vbTextCompare) = 0) Then
            Index = i
        
            Exit For
        End If
    Next i
    
    If (Index >= m_objCount) Then
        Exit Sub
    End If
    
    Select Case (UCase$(m_arrObjs(Index).ObjType))
        Case "BUTTON"
            If (m_arrObjs(Index).obj.Index > 0) Then
                Unload cmd(m_arrObjs(Index).obj.Index)
            Else
                cmd(0).Visible = False
            End If
        
        Case "CHECKBOX"
            If (m_arrObjs(Index).obj.Index > 0) Then
                Unload chk(m_arrObjs(Index).obj.Index)
            Else
                chk(0).Visible = False
            End If
        
        Case "COMBOBOX"
            If (m_arrObjs(Index).obj.Index > 0) Then
                Unload cmb(m_arrObjs(Index).obj.Index)
            Else
                cmb(0).Visible = False
            End If
        
        Case "FRAME"
            If (m_arrObjs(Index).obj.Index > 0) Then
                Unload fra(m_arrObjs(Index).obj.Index)
            Else
                fra(0).Visible = False
            End If
        
        Case "IMAGELIST"
            If (m_arrObjs(Index).obj.Index > 0) Then
                Unload iml(m_arrObjs(Index).obj.Index)
            Else
                iml(0).ListImages.Clear
            End If
        
        Case "LABEL"
            If (m_arrObjs(Index).obj.Index > 0) Then
                Unload lbl(m_arrObjs(Index).obj.Index)
            Else
                lbl(0).Visible = False
            End If
        
        Case "LISTBOX"
            If (m_arrObjs(Index).obj.Index > 0) Then
                Unload lst(m_arrObjs(Index).obj.Index)
            Else
                With lst(0)
                    .Clear
                    .Visible = False
                End With
            End If
        
        Case "LISTVIEW"
            If (m_arrObjs(Index).obj.Index > 0) Then
                Unload lsv(m_arrObjs(Index).obj.Index)
            Else
                With lsv(0)
                    .ListItems.Clear
                    .Visible = False
                End With
            End If
            
        Case "MENU"
        
        Case "OPTIONBUTTON"
            If (m_arrObjs(Index).obj.Index > 0) Then
                Unload opt(m_arrObjs(Index).obj.Index)
            Else
                opt(0).Visible = False
            End If
        
        Case "PICTUREBOX"
            If (m_arrObjs(Index).obj.Index > 0) Then
                Unload pic(m_arrObjs(Index).obj.Index)
            Else
                pic(0).Visible = False
            End If
        
        Case "PROGRESSBAR"
            If (m_arrObjs(Index).obj.Index > 0) Then
                Unload prg(m_arrObjs(Index).obj.Index)
            Else
                With prg(0)
                    .Value = 0
                    .Visible = False
                End With
            End If
        
        Case "RICHTEXTBOX"
            If (m_arrObjs(Index).obj.Index > 0) Then
                Unload rtb(m_arrObjs(Index).obj.Index)
            Else
                With rtb(0)
                    .Text = ""
                    .Visible = False
                End With
            End If
        
        Case "TEXTBOX"
            If (m_arrObjs(Index).obj.Index > 0) Then
                Unload txt(m_arrObjs(Index).obj.Index)
            Else
                With txt(0)
                    .Text = ""
                    .Visible = False
                End With
            End If
            
        Case "TREEVIEW"
            If (m_arrObjs(Index).obj.Index > 0) Then
                Unload trv(m_arrObjs(Index).obj.Index)
            Else
                With trv(0)
                    .Nodes.Clear
                    .Visible = False
                End With
            End If
        
    End Select
    
    Set m_arrObjs(Index).obj = Nothing
    
    If (Index < m_objCount) Then
        For i = Index To ((m_objCount - 1) - 1)
            m_arrObjs(i) = m_arrObjs(i + 1)
        Next i
    End If
    
    If (m_objCount > 1) Then
        ReDim Preserve m_arrObjs(0 To m_objCount - 1)
    Else
        ReDim m_arrObjs(0)
    End If
    
    m_objCount = (m_objCount - 1)
    
    Exit Sub
    
ERROR_HANDLER:
    
    frmChat.AddChat RTBColors.ErrorMessageText, _
        "Error (#" & Err.Number & "): " & Err.Description & " in frmScript::DestroyObjs()."
        
    Resume Next
    
End Sub

Public Function GetObjByName(ByVal ObjName As String) As Object
    Dim i As Integer
    
    For i = 0 To m_objCount - 1
        If (StrComp(m_arrObjs(i).ObjName, ObjName, vbTextCompare) = 0) Then
            Set GetObjByName = m_arrObjs(i).obj

            Exit Function
        End If
    Next i
End Function

Private Function GetScriptObjByIndex(ByVal ObjType As String, ByVal Index As Integer) As scObj
    Dim i As Integer

    For i = 0 To m_objCount - 1
        If (StrComp(ObjType, m_arrObjs(i).ObjType, vbTextCompare) = 0) Then
            If (m_arrObjs(i).obj.Index = Index) Then
                GetScriptObjByIndex = m_arrObjs(i)
                
                Exit For
            End If
        End If
    Next i
End Function

Public Sub ClearObjs()
    On Error GoTo ERROR_HANDLER

    Dim i As Integer
    
    For i = m_objCount - 1 To 0 Step -1
        Select Case (UCase$(m_arrObjs(i).ObjType))
            Case "CHECKBOX"
                chk(m_arrObjs(i).obj.Index).Value = vbUnchecked
                
            Case "COMBOXBOX"
                cmb(m_arrObjs(i).obj.Index).Text = ""
            
            Case "FRAME"
            
            Case "IMAGELIST"
                iml(m_arrObjs(i).obj.Index).ListImages.Clear
            
            Case "LISTBOX"
                lst(m_arrObjs(i).obj.Index).Clear
            
            Case "LISTVIEW"
                lsv(m_arrObjs(i).obj.Index).ListItems.Clear
                
            Case "MENU"
            
            Case "OPTIONBUTTON"
                opt(m_arrObjs(i).obj.Index).Value = False

            Case "PICTUREBOX"
                pic(m_arrObjs(i).obj.Index).Picture = Nothing
            
            Case "PROGRESSBAR"
                prg(m_arrObjs(i).obj.Index).Value = 0
            
            Case "RICHTEXTBOX"
                rtb(m_arrObjs(i).obj.Index).Text = ""
                
                DisableURLDetect m_arrObjs(i).obj.hWnd
                
            Case "TEXTBOX"
                txt(m_arrObjs(i).obj.Index).Text = ""
            
            Case "TREEVIEW"
                trv(m_arrObjs(i).obj.Index).Nodes.Clear
            
        End Select
    Next i

    Exit Sub

ERROR_HANDLER:
    
    frmChat.AddChat RTBColors.ErrorMessageText, _
        "Error (#" & Err.Number & "): " & Err.Description & " in ClearObjs()."
        
    Resume Next
End Sub

Public Sub AddChat(ByVal rtbName As String, ParamArray saElements() As Variant)
    Dim arr() As Variant
    
    arr() = saElements
    
    Call DisplayRichText(GetObjByName(rtbName), arr)
End Sub

'//////////////////////////////////////////////////////
'//Events
'//////////////////////////////////////////////////////

Public Sub Initialize()
    On Error Resume Next
    
    RunInSingle m_sc_module, m_name & "_Initialize"
    RunInSingle m_sc_module, m_name & "_Load"
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    
    RunInSingle m_sc_module, m_name & "_Activate"
End Sub

Private Sub Form_Click()
    On Error Resume Next

    RunInSingle m_sc_module, m_name & "_Click"
End Sub

Private Sub Form_DblClick()
    On Error Resume Next

    RunInSingle m_sc_module, m_name & "_DblClick"
End Sub

Private Sub Form_Deactivate()
    On Error Resume Next

    RunInSingle m_sc_module, m_name & "_Deactivate"
End Sub

Private Sub Form_GotFocus()
    On Error Resume Next

    RunInSingle m_sc_module, m_name & "_GotFocus"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next

    RunInSingle m_sc_module, m_name & "_KeyDown", KeyCode, Shift
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next

    If (RunInSingle(m_sc_module, m_name & "_KeyPress", KeyAscii)) Then
        ' vetoed
        KeyAscii = 0
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next

    RunInSingle m_sc_module, m_name & "_KeyUp", KeyCode, Shift
End Sub

Private Sub Form_Load()
    Me.Icon = frmChat.Icon
End Sub

Private Sub Form_LostFocus()
    On Error Resume Next

    RunInSingle m_sc_module, m_name & "_LostFocus"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next

    RunInSingle m_sc_module, m_name & "_MouseDown", Button, Shift, x, y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next

    RunInSingle m_sc_module, m_name & "_MouseMove", Button, Shift, x, y
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next

    RunInSingle m_sc_module, m_name & "_MouseUp", Button, Shift, x, y
End Sub

Private Sub Form_Paint()
    On Error Resume Next
    
    If (m_hidden = True) Then
        RunInSingle m_sc_module, m_name & "_Load"
        
        m_hidden = False
    End If

    RunInSingle m_sc_module, m_name & "_Paint"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next

    If (RunInSingle(m_sc_module, m_name & "_QueryUnload", UnloadMode)) Then
        ' vetoed
        Cancel = 1
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next

    RunInSingle m_sc_module, m_name & "_Resize"
End Sub

Private Sub Form_Terminate()
    On Error Resume Next

    RunInSingle m_sc_module, m_name & "_Terminate"
End Sub

Public Sub Form_Unload(Cancel As Integer)
    On Error Resume Next

    If (m_hidden = False) Then
        If (RunInSingle(m_sc_module, m_name & "_Unload")) Then
            ' vetoed
            Exit Sub
        End If
        
        Me.Hide
        m_hidden = True
        Cancel = 1
    End If
End Sub

Private Sub cmd_LostFocus(Index As Integer)
    On Error Resume Next

    Dim obj As scObj
    
    obj = GetScriptObjByIndex("Button", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_LostFocus"
End Sub

Private Sub cmd_GotFocus(Index As Integer)
    On Error Resume Next

    Dim obj As scObj
    
    obj = GetScriptObjByIndex("Button", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_GotFocus"
End Sub

Private Sub cmd_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next

    Dim obj As scObj
    
    obj = GetScriptObjByIndex("Button", Index)
    If (RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_KeyPress", KeyAscii)) Then
        ' vetoed
        KeyAscii = 0
    End If
End Sub

Private Sub cmd_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next

    Dim obj As scObj
    
    obj = GetScriptObjByIndex("Button", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_KeyUp", KeyCode, Shift
End Sub

Private Sub cmd_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next

    Dim obj As scObj
    
    obj = GetScriptObjByIndex("Button", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_MouseDown", Button, Shift, x, y
End Sub

Private Sub cmd_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next

    Dim obj As scObj
    
    obj = GetScriptObjByIndex("Button", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_MouseMove", Button, Shift, x, y
End Sub

Private Sub cmd_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next

    Dim obj As scObj
    
    obj = GetScriptObjByIndex("Button", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_MouseUp", Button, Shift, x, y
End Sub

Private Sub cmd_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next

    Dim obj As scObj
    
    obj = GetScriptObjByIndex("Button", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_KeyDown", KeyCode, Shift
End Sub

Private Sub cmd_Click(Index As Integer)
    On Error Resume Next

    Dim obj As scObj
    
    obj = GetScriptObjByIndex("Button", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_Click"
End Sub

Private Sub lbl_Change(Index As Integer)
    On Error Resume Next

    Dim obj As scObj
    
    obj = GetScriptObjByIndex("Label", Index)
    
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_Change"
End Sub

Private Sub lbl_Click(Index As Integer)
    On Error Resume Next

    Dim obj As scObj
    
    obj = GetScriptObjByIndex("Label", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_Click"
End Sub

Private Sub lbl_DblClick(Index As Integer)
    On Error Resume Next

    Dim obj As scObj
    
    obj = GetScriptObjByIndex("Label", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_DblClick"
End Sub

Private Sub lbl_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    
    Dim obj As scObj
    
    obj = GetScriptObjByIndex("Label", Index)
    
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_MouseDown", Button, Shift, x, y
End Sub

Private Sub lbl_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next

    Dim obj As scObj
    
    obj = GetScriptObjByIndex("Label", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_MouseMove", Button, Shift, x, y
End Sub

Private Sub lbl_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next

    Dim obj As scObj
    
    obj = GetScriptObjByIndex("Label", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_MouseUp", Button, Shift, x, y
End Sub

Private Sub lst_Click(Index As Integer)
    On Error Resume Next

    Dim obj As scObj
    
    obj = GetScriptObjByIndex("ListBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_Click"
End Sub

Private Sub lst_DblClick(Index As Integer)
    On Error Resume Next

    Dim obj As scObj
    
    obj = GetScriptObjByIndex("ListBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_DblClick"
End Sub

Private Sub lst_GotFocus(Index As Integer)
    On Error Resume Next

    Dim obj As scObj
    
    obj = GetScriptObjByIndex("ListBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_GotFocus"
End Sub

Private Sub lst_ItemCheck(Index As Integer, Item As Integer)
    On Error Resume Next

    Dim obj As scObj
    
    obj = GetScriptObjByIndex("ListBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_ItemCheck", Item
End Sub

Private Sub lst_ItemClick(Index As Integer, Item As Integer)
    On Error Resume Next

    Dim obj As scObj
    
    obj = GetScriptObjByIndex("ListBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_ItemClick", Item
End Sub

Private Sub lst_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next

    Dim obj As scObj
    
    obj = GetScriptObjByIndex("ListBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_KeyDown", KeyCode, Shift
End Sub

Private Sub lst_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next

    Dim obj As scObj
    
    obj = GetScriptObjByIndex("ListBox", Index)
    If (RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_KeyPress", KeyAscii)) Then
        ' vetoed
        KeyAscii = 0
    End If
End Sub

Private Sub lst_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next

    Dim obj As scObj
    
    obj = GetScriptObjByIndex("ListBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_KeyUp", KeyCode, Shift
End Sub

Private Sub lst_LostFocus(Index As Integer)
    On Error Resume Next

    Dim obj As scObj
    
    obj = GetScriptObjByIndex("ListBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_LostFocus"
End Sub

Private Sub lst_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next

    Dim obj As scObj
    
    obj = GetScriptObjByIndex("ListBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_MouseDown", Button, Shift, x, y
End Sub

Private Sub lst_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next

    Dim obj As scObj
    
    obj = GetScriptObjByIndex("ListBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_MouseMove", Button, Shift, x, y
End Sub

Private Sub lst_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("ListBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_MouseUp", Button, Shift, x, y
End Sub

Private Sub lst_Scroll(Index As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("ListBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_Scroll"
End Sub

Private Sub lsv_ItemClick(Index As Integer, ByVal Item As ListItem)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("ListView", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_ItemClick", Item
End Sub

Private Sub lsv_ColumnClick(Index As Integer, ByVal ColumnHeader As ColumnHeader)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("ListView", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_ColumnClick", ColumnHeader
End Sub

Private Sub lsv_AfterLabelEdit(Index As Integer, Cancel As Integer, NewString As String)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("ListView", Index)
    If (RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_AfterLabelEdit", NewString)) Then
        ' vetoed
        Cancel = 1
    End If
End Sub

Private Sub lsv_BeforeLabelEdit(Index As Integer, Cancel As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("ListView", Index)
    If (RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_BeforeLabelEdit")) Then
        ' vetoed
        Cancel = 1
    End If
End Sub

Private Sub lsv_Click(Index As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("ListView", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_Click"
End Sub

Private Sub lsv_DblClick(Index As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("ListView", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_DblClick"
End Sub

Private Sub lsv_GotFocus(Index As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("ListView", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_GotFocus"
End Sub

Private Sub lsv_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("ListView", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_KeyDown", KeyCode, Shift
End Sub

Private Sub lsv_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("ListView", Index)
    If (RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_KeyPress", KeyAscii)) Then
        ' vetoed
        KeyAscii = 0
    End If
End Sub

Private Sub lsv_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("ListView", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_KeyUp", KeyCode, Shift
End Sub

Private Sub lsv_LostFocus(Index As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("ListView", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_LostFocus"
End Sub

Private Sub lsv_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("ListView", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_MouseDown", Button, Shift, x, y
End Sub

Private Sub lsv_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("ListView", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_MouseMove", Button, Shift, x, y
End Sub

Private Sub lsv_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("ListView", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_MouseUp", Button, Shift, x, y
End Sub

Private Sub opt_Click(Index As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("OptionButton", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_Click"
End Sub

Private Sub opt_DblClick(Index As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("OptionButton", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_DblClick"
End Sub

Private Sub opt_GotFocus(Index As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("OptionButton", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_GotFocus"
End Sub

Private Sub opt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("OptionButton", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_KeyDown", KeyCode, Shift
End Sub

Private Sub opt_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("OptionButton", Index)
    If (RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_KeyPress", KeyAscii)) Then
        ' vetoed
        KeyAscii = 0
    End If
End Sub

Private Sub opt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("OptionButton", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_KeyUp", KeyCode, Shift
End Sub

Private Sub opt_LostFocus(Index As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("OptionButton", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_LostFocus"
End Sub

Private Sub opt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("OptionButton", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_MouseDown", Button, Shift, x, y
End Sub

Private Sub opt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("OptionButton", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_MouseMove", Button, Shift, x, y
End Sub

Private Sub opt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("OptionButton", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_MouseUp", Button, Shift, x, y
End Sub

Private Sub pic_Change(Index As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("PictureBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_Change"
End Sub

Private Sub pic_Click(Index As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("PictureBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_Click"
End Sub

Private Sub pic_DblClick(Index As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("PictureBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_DblClick"
End Sub

Private Sub pic_GotFocus(Index As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("PictureBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_GotFocus"
End Sub

Private Sub pic_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("PictureBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_KeyDown", KeyCode, Shift
End Sub

Private Sub pic_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("PictureBox", Index)
    If (RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_KeyPress", KeyAscii)) Then
        ' vetoed
        KeyAscii = 0
    End If
End Sub

Private Sub pic_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("PictureBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_KeyUp", KeyCode, Shift
End Sub

Private Sub pic_LinkClose(Index As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("PictureBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_LinkClose"
End Sub

Private Sub pic_LinkError(Index As Integer, LinkErr As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("PictureBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_LinkError", LinkErr
End Sub

Private Sub pic_LinkNotify(Index As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("PictureBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_LinkNotify"
End Sub

Private Sub pic_LinkOpen(Index As Integer, Cancel As Integer)

     On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("PictureBox", Index)
    If (RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_LinkOpen")) Then
        ' vetoed
        Cancel = 1
    End If
End Sub

Private Sub pic_LostFocus(Index As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("PictureBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_LostFocus"
End Sub

Private Sub pic_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("PictureBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_MouseDown", Button, Shift, x, y
End Sub

Private Sub pic_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("PictureBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_MouseMove", Button, Shift, x, y
End Sub

Private Sub pic_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
 
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("PictureBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_MouseUp", Button, Shift, x, y
End Sub

Private Sub pic_Paint(Index As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("PictureBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_Paint"
End Sub

Private Sub pic_Resize(Index As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("PictureBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_Resize"
    
End Sub

Private Sub rtb_Change(Index As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("RichTextBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_Change"
End Sub

Private Sub rtb_Click(Index As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("RichTextBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_Click"
End Sub

Private Sub rtb_DblClick(Index As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("RichTextBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_DblClick"
End Sub

Private Sub rtb_GotFocus(Index As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("RichTextBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_GotFocus"
End Sub

Private Sub rtb_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("RichTextBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_KeyDown", KeyCode, Shift
End Sub

Private Sub rtb_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("RichTextBox", Index)
    If (RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_KeyPress", KeyAscii)) Then
        ' vetoed
        KeyAscii = 0
    End If
    
End Sub

Private Sub rtb_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("RichTextBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_KeyUp", KeyCode, Shift
End Sub

Private Sub rtb_LostFocus(Index As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("RichTextBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_LostFocus"
End Sub

Private Sub rtb_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("RichTextBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_MouseDown", Button, Shift, x, y
End Sub

Private Sub rtb_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
 
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("RichTextBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_MouseMove", Button, Shift, x, y
End Sub

Private Sub rtb_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("RichTextBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_MouseUp", Button, Shift, x, y
End Sub

Private Sub rtb_SelChange(Index As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("RichTextBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_SelChange"
End Sub

Private Sub txt_Change(Index As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("TextBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_Change"
End Sub

Private Sub txt_Click(Index As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("TextBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_Click"
End Sub

Private Sub txt_DblClick(Index As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("TextBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_DblClick"
End Sub

Private Sub txt_LostFocus(Index As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("TextBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_LostFocus"
End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("TextBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_MouseDown", Button, Shift, x, y
End Sub

Private Sub txt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("TextBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_MouseMove", Button, Shift, x, y
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("TextBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_MouseUp", Button, Shift, x, y
End Sub

Private Sub txt_GotFocus(Index As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("TextBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_GotFocus"
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("TextBox", Index)
    If (RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_KeyPress", KeyAscii)) Then
        ' vetoed
        KeyAscii = 0
    End If
End Sub

Private Sub txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("TextBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_KeyUp", KeyCode, Shift
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("TextBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_KeyDown", KeyCode, Shift
End Sub

Private Sub chk_Click(Index As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("CheckBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_Click"
End Sub

Private Sub chk_GotFocus(Index As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("CheckBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_GotFocus"
End Sub

Private Sub chk_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("CheckBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_KeyDown", KeyCode, Shift
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("CheckBox", Index)
    If (RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_KeyPress", KeyAscii)) Then
        ' vetoed
        KeyAscii = 0
    End If
End Sub

Private Sub chk_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("CheckBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_KeyUp", KeyCode, Shift
End Sub

Private Sub chk_LostFocus(Index As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("CheckBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_LostFocus"
End Sub

Private Sub chk_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("CheckBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_MouseDown", Button, Shift, x, y
End Sub

Private Sub chk_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("CheckBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_MouseMode", Button, Shift, x, y
End Sub

Private Sub chk_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("CheckBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_MouseUp", Button, Shift, x, y
End Sub

Private Sub cmb_Change(Index As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("ComboBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_Change"
End Sub

Private Sub cmb_Click(Index As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("ComboBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_Click"
End Sub

Private Sub cmb_DblClick(Index As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("ComboBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_DblClick"
End Sub

Private Sub cmb_GotFocus(Index As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("ComboBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_GotFocus"
End Sub

Private Sub cmb_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("ComboBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_KeyDown", KeyCode, Shift
End Sub

Private Sub cmb_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("ComboBox", Index)
    If (RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_KeyPress", KeyAscii)) Then
        ' vetoed
        KeyAscii = 0
    End If
End Sub

Private Sub cmb_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("ComboBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_KeyUp", KeyCode, Shift
End Sub

Private Sub cmb_LostFocus(Index As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("ComboBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_LostFocus"
End Sub

Private Sub cmb_Scroll(Index As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("ComboBox", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_Scroll"
End Sub

Private Sub trv_AfterLabelEdit(Index As Integer, Cancel As Integer, NewString As String)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("TreeView", Index)
    If (RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_AfterLabelEdit", NewString)) Then
        ' vetoed
        Cancel = 1
    End If
End Sub

Private Sub trv_BeforeLabelEdit(Index As Integer, Cancel As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("TreeView", Index)
    If (RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_BeforeLabelEdit")) Then
        ' vetoed
        Cancel = 1
    End If
End Sub

Private Sub trv_Click(Index As Integer)
    
    On Error Resume Next
    
    Dim obj As scObj
    
    obj = GetScriptObjByIndex("TreeView", Index)
    
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_Click"
End Sub

Private Sub trv_Collapse(Index As Integer, ByVal Node As Node)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("TreeView", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_Collapse", Node
    
End Sub

Private Sub trv_DblClick(Index As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("TreeView", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_DblClick"
End Sub

Private Sub trv_Expand(Index As Integer, ByVal Node As Node)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("TreeView", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_Expand", Node
    
End Sub

Private Sub trv_GotFocus(Index As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("TreeView", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_GotFocus"
End Sub

Private Sub trv_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("TreeView", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_KeyDown", KeyCode, Shift
End Sub

Private Sub trv_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("TreeView", Index)
    If (RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_KeyPress", KeyAscii)) Then
        ' vetoed
        KeyAscii = 0
    End If
End Sub

Private Sub trv_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("TreeView", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_KeyUp", KeyCode, Shift
End Sub

Private Sub trv_LostFocus(Index As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("TreeView", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_LostFocus"
End Sub

Private Sub trv_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    
    Dim obj As scObj
    
    obj = GetScriptObjByIndex("TreeView", Index)
    
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_MouseDown", Button, Shift, x, y
End Sub

Private Sub trv_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    
    Dim obj As scObj
    
    obj = GetScriptObjByIndex("TreeView", Index)
    
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_MouseMove", Button, Shift, x, y
End Sub

Private Sub trv_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    
    Dim obj As scObj
    
    obj = GetScriptObjByIndex("TreeView", Index)
    
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_MouseUp", Button, Shift, x, y
End Sub

Private Sub trv_NodeCheck(Index As Integer, ByVal Node As Node)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("TreeView", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_NodeCheck", Node
    
End Sub

Private Sub trv_NodeClick(Index As Integer, ByVal Node As Node)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("TreeView", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_NodeClick", Node
    
End Sub

Private Sub prg_Click(Index As Integer)
    
    On Error Resume Next
    
    Dim obj As scObj
    
    obj = GetScriptObjByIndex("ProgressBar", Index)
    
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_Click"
End Sub

Private Sub prg_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    
    Dim obj As scObj
    
    obj = GetScriptObjByIndex("ProgressBar", Index)
    
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_MouseDown", Button, Shift, x, y
End Sub

Private Sub prg_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    
    Dim obj As scObj
    
    obj = GetScriptObjByIndex("ProgressBar", Index)
    
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_MouseMove", Button, Shift, x, y
End Sub

Private Sub prg_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    
    Dim obj As scObj
    
    obj = GetScriptObjByIndex("ProgressBar", Index)
    
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_MouseUp", Button, Shift, x, y
End Sub

Private Sub fra_Click(Index As Integer)
    
    On Error Resume Next
    
    Dim obj As scObj
    
    obj = GetScriptObjByIndex("Frame", Index)
    
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_Click"
End Sub

Private Sub fra_DblClick(Index As Integer)
    On Error Resume Next

    Dim obj As scObj

    obj = GetScriptObjByIndex("Frame", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_DblClick"
End Sub

Private Sub fra_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    
    Dim obj As scObj
    
    obj = GetScriptObjByIndex("Frame", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_MouseDown", Button, Shift, x, y
End Sub

Private Sub fra_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    
    Dim obj As scObj
    
    obj = GetScriptObjByIndex("Frame", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_MouseMove", Button, Shift, x, y
End Sub

Private Sub fra_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    
    Dim obj As scObj
    
    obj = GetScriptObjByIndex("Frame", Index)
    RunInSingle m_sc_module, m_name & "_" & obj.ObjName & "_MouseUp", Button, Shift, x, y
End Sub

