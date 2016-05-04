Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks
Friend Class frmScript
	Inherits System.Windows.Forms.Form
	
	Private m_name As String
	Private m_sc_module As MSScriptControl.Module
	Private m_arrObjs() As modScripting.scObj
	Private m_objCount As Short
	Private m_hidden As Boolean
	
	'UPGRADE_NOTE: str was upgraded to str_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function setName(ByVal str_Renamed As String) As Object
		
		If (m_name = vbNullString) Then
			m_name = str_Renamed
		End If
	End Function
	
	Public Function getName() As String
		
		getName = m_name
	End Function
	
	Public Function setSCModule(ByRef SCModule As MSScriptControl.Module) As Object
		
		If (m_sc_module Is Nothing) Then
			m_sc_module = SCModule
		End If
	End Function
	
	Public Function GetScriptModule() As MSScriptControl.Module
		
		GetScriptModule = m_sc_module
	End Function
	
	'// 6/22/2009 JSM - Adding wrapper function for MsgBox inside VB6 rather than
	'//                 the scripting control. This keeps the focus on the form.
	' made parameters optional, like the VBs equivalents -Ribose/2009-08-10
	'UPGRADE_NOTE: Text was upgraded to Text_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function ShowMsgBox(ByVal Text_Renamed As String, Optional ByVal opts As MsgBoxStyle = MsgBoxStyle.OKOnly, Optional ByVal Title As String = vbNullString) As MsgBoxResult
		
		ShowMsgBox = MsgBox(Text_Renamed, opts, Title)
	End Function
	
	' wrapper function for InputBox, too!
	' made parameters optional, like the VBs equivalents -Ribose/2009-08-10
	'UPGRADE_NOTE: Default was upgraded to Default_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_NOTE: Text was upgraded to Text_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function ShowInputBox(ByVal Text_Renamed As String, Optional ByVal Title As String = vbNullString, Optional ByVal Default_Renamed As String = vbNullString, Optional ByVal XPos As Short = -1, Optional ByVal YPos As Short = -1) As String
		
		If XPos = -1 And YPos = -1 Then
			ShowInputBox = InputBox(Text_Renamed, Title, Default_Renamed)
		ElseIf XPos = -1 Then 
            ShowInputBox = InputBox(Text_Renamed, Title, Default_Renamed, -1, VB6.TwipsToPixelsY(YPos))
		ElseIf YPos = -1 Then 
			ShowInputBox = InputBox(Text_Renamed, Title, Default_Renamed, VB6.TwipsToPixelsX(XPos))
		Else
			ShowInputBox = InputBox(Text_Renamed, Title, Default_Renamed, VB6.TwipsToPixelsX(XPos), VB6.TwipsToPixelsY(YPos))
		End If
		
	End Function
	
	Public Sub DrawLine(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, Optional ByVal Color As Integer = -1, Optional ByVal DrawRect As Boolean = False, Optional ByVal FillRect As Boolean = False)
		
		If Color = -1 Then
			If DrawRect Then
				If FillRect Then
					'UPGRADE_ISSUE: Form method frmScript.Line was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					Me.Line (x1, y1) - (x2, y2), BF
				Else
					'UPGRADE_ISSUE: Form method frmScript.Line was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					Me.Line (x1, y1) - (x2, y2), B
				End If
			Else
				'UPGRADE_ISSUE: Form method frmScript.Line was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				Me.Line (x1, y1) - (x2, y2)
			End If
		Else
			If DrawRect Then
				If FillRect Then
					'UPGRADE_ISSUE: Form method frmScript.Line was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					Me.Line (x1, y1) - (x2, y2), Color, BF
				Else
					'UPGRADE_ISSUE: Form method frmScript.Line was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					Me.Line (x1, y1) - (x2, y2), Color, B
				End If
			Else
				'UPGRADE_ISSUE: Form method frmScript.Line was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				Me.Line (x1, y1) - (x2, y2), Color
			End If
		End If
	End Sub
	
	'Public Function Objects(objIndex As Integer) As scObj
	'
	'    Objects = m_arrObjs(objIndex)
	'
	'End Function
	
	Public Function ObjCount(Optional ByRef ObjType As String = "") As Short
		Dim i As Short
		If (ObjType <> vbNullString) Then
			For i = 0 To m_objCount - 1
				If (StrComp(ObjType, m_arrObjs(i).ObjType, CompareMethod.Text) = 0) Then
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
		
		'UPGRADE_NOTE: Object CreateObj may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		CreateObj = Nothing
		If (Not ValidObjectName(ObjName)) Then Exit Function
		
		' redefine array size & check for duplicate controls
		Dim i As Short ' loop counter variable
		If (m_objCount) Then
			
			For i = 0 To m_objCount - 1
				If (StrComp(m_arrObjs(i).ObjType, ObjType, CompareMethod.Text) = 0) Then
					If (StrComp(m_arrObjs(i).ObjName, ObjName, CompareMethod.Text) = 0) Then
						CreateObj = m_arrObjs(i).obj
						
						Exit Function
					End If
				End If
			Next i
			
			ReDim Preserve m_arrObjs(m_objCount)
		Else
			ReDim m_arrObjs(0)
		End If
		
		Select Case (UCase(ObjType))
			Case "BUTTON"
				If (ObjCount(ObjType) > 0) Then
					cmd.Load(ObjCount(ObjType))
				End If
				
				obj.obj = cmd(ObjCount(ObjType))
				
			Case "CHECKBOX"
				If (ObjCount(ObjType) > 0) Then
					chk.Load(ObjCount(ObjType))
				End If
				
				obj.obj = chk(ObjCount(ObjType))
				
			Case "COMBOBOX"
				If (ObjCount(ObjType) > 0) Then
					cmb.Load(ObjCount(ObjType))
				End If
				
				obj.obj = cmb(ObjCount(ObjType))
				
			Case "FRAME"
				If (ObjCount(ObjType) > 0) Then
					fra.Load(ObjCount(ObjType))
				End If
				
				obj.obj = fra(ObjCount(ObjType))
				
			Case "IMAGELIST"
				If (ObjCount(ObjType) > 0) Then
					iml.Load(ObjCount(ObjType))
				End If
				
				obj.obj = iml(ObjCount(ObjType))
				
			Case "LABEL"
				If (ObjCount(ObjType) > 0) Then
					lbl.Load(ObjCount(ObjType))
				End If
				
				obj.obj = lbl(ObjCount(ObjType))
				
			Case "LISTBOX"
				If (ObjCount(ObjType) > 0) Then
					lst.Load(ObjCount(ObjType))
				End If
				
				obj.obj = lst(ObjCount(ObjType))
				
			Case "LISTVIEW"
				If (ObjCount(ObjType) > 0) Then
					lsv.Load(ObjCount(ObjType))
				End If
				
				obj.obj = lsv(ObjCount(ObjType))
				
			Case "MENU"
				obj.obj = New clsMenuObj
				
				'UPGRADE_WARNING: Couldn't resolve default property of object obj.obj.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				obj.obj.Name = getName() & "_" & ObjName
				
				'UPGRADE_WARNING: Couldn't resolve default property of object obj.obj.Parent. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object Me. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				obj.obj.Parent = Me
				
				DynamicMenus.Add(obj.obj)
				
			Case "OPTIONBUTTON"
				If (ObjCount(ObjType) > 0) Then
					opt.Load(ObjCount(ObjType))
				End If
				
				obj.obj = opt(ObjCount(ObjType))
				
			Case "PICTUREBOX"
				If (ObjCount(ObjType) > 0) Then
					pic.Load(ObjCount(ObjType))
				End If
				
				obj.obj = pic(ObjCount(ObjType))
				
			Case "PROGRESSBAR"
				If (ObjCount(ObjType) > 0) Then
					prg.Load(ObjCount(ObjType))
				End If
				
				obj.obj = prg(ObjCount(ObjType))
				
			Case "RICHTEXTBOX"
				If (ObjCount(ObjType) > 0) Then
					rtb.Load(ObjCount(ObjType))
				End If
				
				obj.obj = rtb(ObjCount(ObjType))
				
				'UPGRADE_WARNING: Couldn't resolve default property of object obj.obj.hWnd. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				EnableURLDetect(obj.obj.hWnd)
				
			Case "TEXTBOX"
				If (ObjCount(ObjType) > 0) Then
					txt.Load(ObjCount(ObjType))
				End If
				
				obj.obj = txt(ObjCount(ObjType))
				
			Case "TREEVIEW"
				If (ObjCount(ObjType) > 0) Then
					trv.Load(ObjCount(ObjType))
				End If
				
				obj.obj = trv(ObjCount(ObjType))
				
		End Select
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj.obj.Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj.obj.Visible = True
		
		' store our module name & type
		obj.ObjName = ObjName
		obj.ObjType = ObjType
		
		' store object
		'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs(m_objCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_arrObjs(m_objCount) = obj
		
		' increment object counter
		m_objCount = (m_objCount + 1)
		
		' return object
		CreateObj = obj.obj
	End Function
	
	Public Sub DestroyObjs()
		
		On Error GoTo ERROR_HANDLER
		
		Dim i As Short
		
		For i = m_objCount - 1 To 0 Step -1
			DestroyObj(m_arrObjs(i).ObjName)
		Next i
		
		Exit Sub
		
ERROR_HANDLER: 
		
		frmChat.AddChat(RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & " in frmScript::DestroyObjs().")
		
		Resume Next
		
	End Sub
	
	Public Sub DestroyObj(ByVal ObjName As String)
		
		On Error GoTo ERROR_HANDLER
		
		Dim i As Short
		Dim Index As Short
		
		If (m_objCount = 0) Then
			Exit Sub
		End If
		
		Index = m_objCount
		
		For i = 0 To m_objCount - 1
			If (StrComp(m_arrObjs(i).ObjName, ObjName, CompareMethod.Text) = 0) Then
				Index = i
				
				Exit For
			End If
		Next i
		
		If (Index >= m_objCount) Then
			Exit Sub
		End If
		
		Select Case (UCase(m_arrObjs(Index).ObjType))
			Case "BUTTON"
				'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs(Index).obj.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (m_arrObjs(Index).obj.Index > 0) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs().obj.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					cmd.Unload(m_arrObjs(Index).obj.Index)
				Else
					cmd(0).Visible = False
				End If
				
			Case "CHECKBOX"
				'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs(Index).obj.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (m_arrObjs(Index).obj.Index > 0) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs().obj.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					chk.Unload(m_arrObjs(Index).obj.Index)
				Else
					chk(0).Visible = False
				End If
				
			Case "COMBOBOX"
				'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs(Index).obj.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (m_arrObjs(Index).obj.Index > 0) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs().obj.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					cmb.Unload(m_arrObjs(Index).obj.Index)
				Else
					cmb(0).Visible = False
				End If
				
			Case "FRAME"
				'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs(Index).obj.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (m_arrObjs(Index).obj.Index > 0) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs().obj.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					fra.Unload(m_arrObjs(Index).obj.Index)
				Else
					fra(0).Visible = False
				End If
				
			Case "IMAGELIST"
				'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs(Index).obj.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (m_arrObjs(Index).obj.Index > 0) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs().obj.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					iml.Unload(m_arrObjs(Index).obj.Index)
				Else
					iml(0).Images.Clear()
				End If
				
			Case "LABEL"
				'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs(Index).obj.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (m_arrObjs(Index).obj.Index > 0) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs().obj.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					lbl.Unload(m_arrObjs(Index).obj.Index)
				Else
					lbl(0).Visible = False
				End If
				
			Case "LISTBOX"
				'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs(Index).obj.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (m_arrObjs(Index).obj.Index > 0) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs().obj.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					lst.Unload(m_arrObjs(Index).obj.Index)
				Else
					With lst(0)
						.Items.Clear()
						.Visible = False
					End With
				End If
				
			Case "LISTVIEW"
				'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs(Index).obj.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (m_arrObjs(Index).obj.Index > 0) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs().obj.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					lsv.Unload(m_arrObjs(Index).obj.Index)
				Else
					With lsv(0)
						.Items.Clear()
						.Visible = False
					End With
				End If
				
			Case "MENU"
				
			Case "OPTIONBUTTON"
				'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs(Index).obj.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (m_arrObjs(Index).obj.Index > 0) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs().obj.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					opt.Unload(m_arrObjs(Index).obj.Index)
				Else
					opt(0).Visible = False
				End If
				
			Case "PICTUREBOX"
				'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs(Index).obj.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (m_arrObjs(Index).obj.Index > 0) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs().obj.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					pic.Unload(m_arrObjs(Index).obj.Index)
				Else
					pic(0).Visible = False
				End If
				
			Case "PROGRESSBAR"
				'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs(Index).obj.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (m_arrObjs(Index).obj.Index > 0) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs().obj.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					prg.Unload(m_arrObjs(Index).obj.Index)
				Else
					With prg(0)
						.Value = 0
						.Visible = False
					End With
				End If
				
			Case "RICHTEXTBOX"
				'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs(Index).obj.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (m_arrObjs(Index).obj.Index > 0) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs().obj.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					rtb.Unload(m_arrObjs(Index).obj.Index)
				Else
					With rtb(0)
						.Text = ""
						.Visible = False
					End With
				End If
				
			Case "TEXTBOX"
				'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs(Index).obj.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (m_arrObjs(Index).obj.Index > 0) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs().obj.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					txt.Unload(m_arrObjs(Index).obj.Index)
				Else
					With txt(0)
						.Text = ""
						.Visible = False
					End With
				End If
				
			Case "TREEVIEW"
				'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs(Index).obj.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (m_arrObjs(Index).obj.Index > 0) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs().obj.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					trv.Unload(m_arrObjs(Index).obj.Index)
				Else
					With trv(0)
						.Nodes.Clear()
						.Visible = False
					End With
				End If
				
		End Select
		
		'UPGRADE_NOTE: Object m_arrObjs().obj may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_arrObjs(Index).obj = Nothing
		
		If (Index < m_objCount) Then
			For i = Index To ((m_objCount - 1) - 1)
				'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				m_arrObjs(i) = m_arrObjs(i + 1)
			Next i
		End If
		
		If (m_objCount > 1) Then
			ReDim Preserve m_arrObjs(m_objCount - 1)
		Else
			ReDim m_arrObjs(0)
		End If
		
		m_objCount = (m_objCount - 1)
		
		Exit Sub
		
ERROR_HANDLER: 
		
		frmChat.AddChat(RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & " in frmScript::DestroyObjs().")
		
		Resume Next
		
	End Sub
	
	Public Function GetObjByName(ByVal ObjName As String) As Object
		Dim i As Short
		
		For i = 0 To m_objCount - 1
			If (StrComp(m_arrObjs(i).ObjName, ObjName, CompareMethod.Text) = 0) Then
				GetObjByName = m_arrObjs(i).obj
				
				Exit Function
			End If
		Next i
	End Function
	
	Private Function GetScriptObjByIndex(ByVal ObjType As String, ByVal Index As Short) As scObj
		Dim i As Short
		
		For i = 0 To m_objCount - 1
			If (StrComp(ObjType, m_arrObjs(i).ObjType, CompareMethod.Text) = 0) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs(i).obj.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (m_arrObjs(i).obj.Index = Index) Then
					GetScriptObjByIndex = m_arrObjs(i)
					
					Exit For
				End If
			End If
		Next i
	End Function
	
	Public Sub ClearObjs()
		On Error GoTo ERROR_HANDLER
		
		Dim i As Short
		
		For i = m_objCount - 1 To 0 Step -1
			Select Case (UCase(m_arrObjs(i).ObjType))
				Case "CHECKBOX"
					'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs().obj.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					chk(m_arrObjs(i).obj.Index).CheckState = System.Windows.Forms.CheckState.Unchecked
					
				Case "COMBOXBOX"
					'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs().obj.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					cmb(m_arrObjs(i).obj.Index).Text = ""
					
				Case "FRAME"
					
				Case "IMAGELIST"
					'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs().obj.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					iml(m_arrObjs(i).obj.Index).Images.Clear()
					
				Case "LISTBOX"
					'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs().obj.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					lst(m_arrObjs(i).obj.Index).Items.Clear()
					
				Case "LISTVIEW"
					'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs().obj.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					lsv(m_arrObjs(i).obj.Index).Items.Clear()
					
				Case "MENU"
					
				Case "OPTIONBUTTON"
					'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs().obj.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					opt(m_arrObjs(i).obj.Index).Checked = False
					
				Case "PICTUREBOX"
					'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs().obj.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_NOTE: Object pic().Picture may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					pic(m_arrObjs(i).obj.Index).Image = Nothing
					
				Case "PROGRESSBAR"
					'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs().obj.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					prg(m_arrObjs(i).obj.Index).Value = 0
					
				Case "RICHTEXTBOX"
					'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs().obj.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					rtb(m_arrObjs(i).obj.Index).Text = ""
					
					'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs().obj.hWnd. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					DisableURLDetect(m_arrObjs(i).obj.hWnd)
					
				Case "TEXTBOX"
					'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs().obj.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					txt(m_arrObjs(i).obj.Index).Text = ""
					
				Case "TREEVIEW"
					'UPGRADE_WARNING: Couldn't resolve default property of object m_arrObjs().obj.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					trv(m_arrObjs(i).obj.Index).Nodes.Clear()
					
			End Select
		Next i
		
		Exit Sub
		
ERROR_HANDLER: 
		
		frmChat.AddChat(RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & " in ClearObjs().")
		
		Resume Next
	End Sub
	
	'UPGRADE_WARNING: ParamArray saElements was changed from ByRef to ByVal. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="93C6A0DC-8C99-429A-8696-35FC4DCEFCCC"'
	Public Sub AddChat(ByVal rtbName As String, ParamArray ByVal saElements() As Object)
		Dim arr() As Object
		
		arr = VB6.CopyArray(saElements)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object GetObjByName(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Call DisplayRichText(GetObjByName(rtbName), arr)
	End Sub
	
	'//////////////////////////////////////////////////////
	'//Events
	'//////////////////////////////////////////////////////
	
	Public Sub Initialize()
		On Error Resume Next
		
		RunInSingle(m_sc_module, m_name & "_Initialize")
		RunInSingle(m_sc_module, m_name & "_Load")
	End Sub
	
	'UPGRADE_WARNING: Form event frmScript.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmScript_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		On Error Resume Next
		
		RunInSingle(m_sc_module, m_name & "_Activate")
	End Sub
	
	Private Sub frmScript_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Click
		On Error Resume Next
		
		RunInSingle(m_sc_module, m_name & "_Click")
	End Sub
	
	Private Sub frmScript_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.DoubleClick
		On Error Resume Next
		
		RunInSingle(m_sc_module, m_name & "_DblClick")
	End Sub
	
	'UPGRADE_WARNING: Form event frmScript.Deactivate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmScript_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
		On Error Resume Next
		
		RunInSingle(m_sc_module, m_name & "_Deactivate")
	End Sub
	
	Private Sub frmScript_GotFocus(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.GotFocus
		On Error Resume Next
		
		RunInSingle(m_sc_module, m_name & "_GotFocus")
	End Sub
	
	Private Sub frmScript_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		On Error Resume Next
		
		RunInSingle(m_sc_module, m_name & "_KeyDown", KeyCode, Shift)
	End Sub
	
	Private Sub frmScript_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		On Error Resume Next
		
		If (RunInSingle(m_sc_module, m_name & "_KeyPress", KeyAscii)) Then
			' vetoed
			KeyAscii = 0
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub frmScript_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyUp
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		On Error Resume Next
		
		RunInSingle(m_sc_module, m_name & "_KeyUp", KeyCode, Shift)
	End Sub
	
	Private Sub frmScript_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Me.Icon = frmChat.Icon
	End Sub
	
	Private Sub frmScript_LostFocus(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.LostFocus
		On Error Resume Next
		
		RunInSingle(m_sc_module, m_name & "_LostFocus")
	End Sub
	
	Private Sub frmScript_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		On Error Resume Next
		
		RunInSingle(m_sc_module, m_name & "_MouseDown", Button, Shift, x, y)
	End Sub
	
	Private Sub frmScript_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		On Error Resume Next
		
		RunInSingle(m_sc_module, m_name & "_MouseMove", Button, Shift, x, y)
	End Sub
	
	Private Sub frmScript_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		On Error Resume Next
		
		RunInSingle(m_sc_module, m_name & "_MouseUp", Button, Shift, x, y)
	End Sub
	
	Private Sub frmScript_Paint(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
		On Error Resume Next
		
		If (m_hidden = True) Then
			RunInSingle(m_sc_module, m_name & "_Load")
			
			m_hidden = False
		End If
		
		RunInSingle(m_sc_module, m_name & "_Paint")
	End Sub
	
	Private Sub frmScript_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
		Dim Cancel As Boolean = eventArgs.Cancel
		Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
		On Error Resume Next
		
		If (RunInSingle(m_sc_module, m_name & "_QueryUnload", UnloadMode)) Then
			' vetoed
			Cancel = 1
		End If
		eventArgs.Cancel = Cancel
	End Sub
	
	'UPGRADE_WARNING: Event frmScript.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub frmScript_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		On Error Resume Next
		
		RunInSingle(m_sc_module, m_name & "_Resize")
	End Sub
	
	'UPGRADE_NOTE: Form_Terminate was upgraded to Form_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_WARNING: frmScript event Form.Terminate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub Form_Terminate_Renamed()
		On Error Resume Next
		
		RunInSingle(m_sc_module, m_name & "_Terminate")
	End Sub
	
	Public Sub frmScript_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		On Error Resume Next
		
		If (m_hidden = False) Then
			If (RunInSingle(m_sc_module, m_name & "_Unload")) Then
				' vetoed
				Exit Sub
			End If
			
			Me.Hide()
			m_hidden = True
			'UPGRADE_ISSUE: Event parameter Cancel was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FB723E3C-1C06-4D2B-B083-E6CD0D334DA8"'
			Cancel = 1
		End If
	End Sub
	
	Private Sub cmd_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmd.Leave
		Dim Index As Short = cmd.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("Button", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_LostFocus")
	End Sub
	
	Private Sub cmd_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmd.Enter
		Dim Index As Short = cmd.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("Button", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_GotFocus")
	End Sub
	
	Private Sub cmd_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles cmd.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = cmd.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("Button", Index)
		If (RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_KeyPress", KeyAscii)) Then
			' vetoed
			KeyAscii = 0
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub cmd_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles cmd.KeyUp
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Dim Index As Short = cmd.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("Button", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_KeyUp", KeyCode, Shift)
	End Sub
	
	Private Sub cmd_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles cmd.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim Index As Short = cmd.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("Button", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_MouseDown", Button, Shift, x, y)
	End Sub
	
	Private Sub cmd_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles cmd.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim Index As Short = cmd.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("Button", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_MouseMove", Button, Shift, x, y)
	End Sub
	
	Private Sub cmd_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles cmd.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim Index As Short = cmd.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("Button", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_MouseUp", Button, Shift, x, y)
	End Sub
	
	Private Sub cmd_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles cmd.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Dim Index As Short = cmd.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("Button", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_KeyDown", KeyCode, Shift)
	End Sub
	
	Private Sub cmd_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmd.Click
		Dim Index As Short = cmd.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("Button", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_Click")
	End Sub
	
	'UPGRADE_ISSUE: Label event lbl.Change was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub lbl_Change(ByRef Index As Short)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("Label", Index)
		
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_Change")
	End Sub
	
	Private Sub lbl_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lbl.Click
		Dim Index As Short = lbl.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("Label", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_Click")
	End Sub
	
	Private Sub lbl_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lbl.DoubleClick
		Dim Index As Short = lbl.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("Label", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_DblClick")
	End Sub
	
	Private Sub lbl_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles lbl.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim Index As Short = lbl.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("Label", Index)
		
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_MouseDown", Button, Shift, x, y)
	End Sub
	
	Private Sub lbl_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles lbl.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim Index As Short = lbl.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("Label", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_MouseMove", Button, Shift, x, y)
	End Sub
	
	Private Sub lbl_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles lbl.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim Index As Short = lbl.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("Label", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_MouseUp", Button, Shift, x, y)
	End Sub
	
	'UPGRADE_WARNING: Event lst.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub lst_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lst.SelectedIndexChanged
		Dim Index As Short = lst.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("ListBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_Click")
	End Sub
	
	Private Sub lst_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lst.DoubleClick
		Dim Index As Short = lst.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("ListBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_DblClick")
	End Sub
	
	Private Sub lst_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lst.Enter
		Dim Index As Short = lst.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("ListBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_GotFocus")
	End Sub
	
	'UPGRADE_ISSUE: ListBox event lst.ItemCheck was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub lst_ItemCheck(ByRef Index As Short, ByRef Item As Short)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("ListBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_ItemCheck", Item)
	End Sub
	
	Private Sub lst_ItemClick(ByRef Index As Short, ByRef Item As Short)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("ListBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_ItemClick", Item)
	End Sub
	
	Private Sub lst_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles lst.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Dim Index As Short = lst.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("ListBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_KeyDown", KeyCode, Shift)
	End Sub
	
	Private Sub lst_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles lst.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = lst.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("ListBox", Index)
		If (RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_KeyPress", KeyAscii)) Then
			' vetoed
			KeyAscii = 0
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub lst_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles lst.KeyUp
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Dim Index As Short = lst.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("ListBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_KeyUp", KeyCode, Shift)
	End Sub
	
	Private Sub lst_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lst.Leave
		Dim Index As Short = lst.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("ListBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_LostFocus")
	End Sub
	
	Private Sub lst_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles lst.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim Index As Short = lst.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("ListBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_MouseDown", Button, Shift, x, y)
	End Sub
	
	Private Sub lst_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles lst.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim Index As Short = lst.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("ListBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_MouseMove", Button, Shift, x, y)
	End Sub
	
	Private Sub lst_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles lst.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim Index As Short = lst.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("ListBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_MouseUp", Button, Shift, x, y)
	End Sub
	
	'UPGRADE_ISSUE: ListBox event lst.Scroll was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub lst_Scroll(ByRef Index As Short)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("ListBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_Scroll")
	End Sub
	
	'UPGRADE_ISSUE: MSComctlLib.ListView event lsv.ItemClick was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub lsv_ItemClick(ByRef Index As Short, ByVal Item As System.Windows.Forms.ListViewItem)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("ListView", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_ItemClick", Item)
	End Sub
	
	Private Sub lsv_ColumnClick(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.ColumnClickEventArgs) Handles lsv.ColumnClick
		Dim Index As Short = lsv.GetIndex(eventSender)
		Dim ColumnHeader As System.Windows.Forms.ColumnHeader = lsv(Index).Columns(eventArgs.Column)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("ListView", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_ColumnClick", ColumnHeader)
	End Sub
	
	Private Sub lsv_AfterLabelEdit(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.LabelEditEventArgs) Handles lsv.AfterLabelEdit
		Dim Cancel As Boolean = eventArgs.CancelEdit
		Dim NewString As String = eventArgs.Label
		Dim Index As Short = lsv.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("ListView", Index)
		If (RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_AfterLabelEdit", NewString)) Then
			' vetoed
			Cancel = 1
		End If
	End Sub
	
	Private Sub lsv_BeforeLabelEdit(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.LabelEditEventArgs) Handles lsv.BeforeLabelEdit
		Dim Cancel As Boolean = eventArgs.CancelEdit
		Dim Index As Short = lsv.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("ListView", Index)
		If (RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_BeforeLabelEdit")) Then
			' vetoed
			Cancel = 1
		End If
	End Sub
	
	Private Sub lsv_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lsv.Click
		Dim Index As Short = lsv.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("ListView", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_Click")
	End Sub
	
	Private Sub lsv_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lsv.DoubleClick
		Dim Index As Short = lsv.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("ListView", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_DblClick")
	End Sub
	
	Private Sub lsv_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lsv.Enter
		Dim Index As Short = lsv.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("ListView", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_GotFocus")
	End Sub
	
	Private Sub lsv_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles lsv.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Dim Index As Short = lsv.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("ListView", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_KeyDown", KeyCode, Shift)
	End Sub
	
	Private Sub lsv_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles lsv.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = lsv.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("ListView", Index)
		If (RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_KeyPress", KeyAscii)) Then
			' vetoed
			KeyAscii = 0
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub lsv_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles lsv.KeyUp
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Dim Index As Short = lsv.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("ListView", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_KeyUp", KeyCode, Shift)
	End Sub
	
	Private Sub lsv_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lsv.Leave
		Dim Index As Short = lsv.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("ListView", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_LostFocus")
	End Sub
	
	Private Sub lsv_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles lsv.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim Index As Short = lsv.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("ListView", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_MouseDown", Button, Shift, x, y)
	End Sub
	
	Private Sub lsv_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles lsv.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim Index As Short = lsv.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("ListView", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_MouseMove", Button, Shift, x, y)
	End Sub
	
	Private Sub lsv_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles lsv.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim Index As Short = lsv.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("ListView", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_MouseUp", Button, Shift, x, y)
	End Sub
	
	'UPGRADE_WARNING: Event opt.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub opt_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles opt.CheckedChanged
		If eventSender.Checked Then
			Dim Index As Short = opt.GetIndex(eventSender)
			On Error Resume Next
			
			Dim obj As scObj
			
			'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			obj = GetScriptObjByIndex("OptionButton", Index)
			RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_Click")
		End If
	End Sub
	
	'UPGRADE_ISSUE: OptionButton event opt.DblClick was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub opt_DblClick(ByRef Index As Short)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("OptionButton", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_DblClick")
	End Sub
	
	Private Sub opt_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles opt.Enter
		Dim Index As Short = opt.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("OptionButton", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_GotFocus")
	End Sub
	
	Private Sub opt_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles opt.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Dim Index As Short = opt.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("OptionButton", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_KeyDown", KeyCode, Shift)
	End Sub
	
	Private Sub opt_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles opt.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = opt.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("OptionButton", Index)
		If (RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_KeyPress", KeyAscii)) Then
			' vetoed
			KeyAscii = 0
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub opt_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles opt.KeyUp
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Dim Index As Short = opt.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("OptionButton", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_KeyUp", KeyCode, Shift)
	End Sub
	
	Private Sub opt_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles opt.Leave
		Dim Index As Short = opt.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("OptionButton", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_LostFocus")
	End Sub
	
	Private Sub opt_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles opt.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim Index As Short = opt.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("OptionButton", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_MouseDown", Button, Shift, x, y)
	End Sub
	
	Private Sub opt_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles opt.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim Index As Short = opt.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("OptionButton", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_MouseMove", Button, Shift, x, y)
	End Sub
	
	Private Sub opt_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles opt.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim Index As Short = opt.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("OptionButton", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_MouseUp", Button, Shift, x, y)
	End Sub
	
	'UPGRADE_ISSUE: PictureBox event pic.Change was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub pic_Change(ByRef Index As Short)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("PictureBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_Change")
	End Sub
	
	Private Sub pic_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles pic.Click
		Dim Index As Short = pic.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("PictureBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_Click")
	End Sub
	
	Private Sub pic_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles pic.DoubleClick
		Dim Index As Short = pic.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("PictureBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_DblClick")
	End Sub
	
	'UPGRADE_ISSUE: PictureBox event pic.GotFocus was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub pic_GotFocus(ByRef Index As Short)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("PictureBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_GotFocus")
	End Sub
	
	'UPGRADE_ISSUE: PictureBox event pic.KeyDown was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub pic_KeyDown(ByRef Index As Short, ByRef KeyCode As Short, ByRef Shift As Short)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("PictureBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_KeyDown", KeyCode, Shift)
	End Sub
	
	'UPGRADE_ISSUE: PictureBox event pic.KeyPress was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub pic_KeyPress(ByRef Index As Short, ByRef KeyAscii As Short)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("PictureBox", Index)
		If (RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_KeyPress", KeyAscii)) Then
			' vetoed
			KeyAscii = 0
		End If
	End Sub
	
	'UPGRADE_ISSUE: PictureBox event pic.KeyUp was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub pic_KeyUp(ByRef Index As Short, ByRef KeyCode As Short, ByRef Shift As Short)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("PictureBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_KeyUp", KeyCode, Shift)
	End Sub
	
	'UPGRADE_ISSUE: PictureBox event pic.LinkClose was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub pic_LinkClose(ByRef Index As Short)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("PictureBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_LinkClose")
	End Sub
	
	'UPGRADE_ISSUE: PictureBox event pic.LinkError was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub pic_LinkError(ByRef Index As Short, ByRef LinkErr As Short)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("PictureBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_LinkError", LinkErr)
	End Sub
	
	'UPGRADE_ISSUE: PictureBox event pic.LinkNotify was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub pic_LinkNotify(ByRef Index As Short)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("PictureBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_LinkNotify")
	End Sub
	
	'UPGRADE_ISSUE: PictureBox event pic.LinkOpen was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub pic_LinkOpen(ByRef Index As Short, ByRef Cancel As Short)
		
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("PictureBox", Index)
		If (RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_LinkOpen")) Then
			' vetoed
			Cancel = 1
		End If
	End Sub
	
	'UPGRADE_ISSUE: PictureBox event pic.LostFocus was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub pic_LostFocus(ByRef Index As Short)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("PictureBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_LostFocus")
	End Sub
	
	Private Sub pic_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles pic.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim Index As Short = pic.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("PictureBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_MouseDown", Button, Shift, x, y)
	End Sub
	
	Private Sub pic_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles pic.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim Index As Short = pic.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("PictureBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_MouseMove", Button, Shift, x, y)
	End Sub
	
	Private Sub pic_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles pic.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim Index As Short = pic.GetIndex(eventSender)
		
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("PictureBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_MouseUp", Button, Shift, x, y)
	End Sub
	
	Private Sub pic_Paint(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.PaintEventArgs) Handles pic.Paint
		Dim Index As Short = pic.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("PictureBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_Paint")
	End Sub
	
	Private Sub pic_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles pic.Resize
		Dim Index As Short = pic.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("PictureBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_Resize")
		
	End Sub
	
	Private Sub rtb_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles rtb.TextChanged
		Dim Index As Short = rtb.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("RichTextBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_Change")
	End Sub
	
	Private Sub rtb_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles rtb.Click
		Dim Index As Short = rtb.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("RichTextBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_Click")
	End Sub
	
	Private Sub rtb_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles rtb.DoubleClick
		Dim Index As Short = rtb.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("RichTextBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_DblClick")
	End Sub
	
	Private Sub rtb_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles rtb.Enter
		Dim Index As Short = rtb.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("RichTextBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_GotFocus")
	End Sub
	
	Private Sub rtb_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles rtb.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Dim Index As Short = rtb.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("RichTextBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_KeyDown", KeyCode, Shift)
	End Sub
	
	Private Sub rtb_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles rtb.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = rtb.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("RichTextBox", Index)
		If (RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_KeyPress", KeyAscii)) Then
			' vetoed
			KeyAscii = 0
		End If
		
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub rtb_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles rtb.KeyUp
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Dim Index As Short = rtb.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("RichTextBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_KeyUp", KeyCode, Shift)
	End Sub
	
	Private Sub rtb_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles rtb.Leave
		Dim Index As Short = rtb.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("RichTextBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_LostFocus")
	End Sub
	
	Private Sub rtb_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles rtb.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim Index As Short = rtb.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("RichTextBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_MouseDown", Button, Shift, x, y)
	End Sub
	
	Private Sub rtb_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles rtb.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim Index As Short = rtb.GetIndex(eventSender)
		
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("RichTextBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_MouseMove", Button, Shift, x, y)
	End Sub
	
	Private Sub rtb_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles rtb.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim Index As Short = rtb.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("RichTextBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_MouseUp", Button, Shift, x, y)
	End Sub
	
	Private Sub rtb_SelectionChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles rtb.SelectionChanged
		Dim Index As Short = rtb.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("RichTextBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_SelChange")
	End Sub
	
	'UPGRADE_WARNING: Event txt.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txt_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txt.TextChanged
		Dim Index As Short = txt.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("TextBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_Change")
	End Sub
	
	Private Sub txt_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txt.Click
		Dim Index As Short = txt.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("TextBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_Click")
	End Sub
	
	Private Sub txt_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txt.DoubleClick
		Dim Index As Short = txt.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("TextBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_DblClick")
	End Sub
	
	Private Sub txt_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txt.Leave
		Dim Index As Short = txt.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("TextBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_LostFocus")
	End Sub
	
	Private Sub txt_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles txt.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim Index As Short = txt.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("TextBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_MouseDown", Button, Shift, x, y)
	End Sub
	
	Private Sub txt_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles txt.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim Index As Short = txt.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("TextBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_MouseMove", Button, Shift, x, y)
	End Sub
	
	Private Sub txt_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles txt.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim Index As Short = txt.GetIndex(eventSender)
		
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("TextBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_MouseUp", Button, Shift, x, y)
	End Sub
	
	Private Sub txt_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txt.Enter
		Dim Index As Short = txt.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("TextBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_GotFocus")
	End Sub
	
	Private Sub txt_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txt.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = txt.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("TextBox", Index)
		If (RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_KeyPress", KeyAscii)) Then
			' vetoed
			KeyAscii = 0
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub txt_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txt.KeyUp
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Dim Index As Short = txt.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("TextBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_KeyUp", KeyCode, Shift)
	End Sub
	
	Private Sub txt_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txt.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Dim Index As Short = txt.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("TextBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_KeyDown", KeyCode, Shift)
	End Sub
	
	'UPGRADE_WARNING: Event chk.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chk_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chk.CheckStateChanged
		Dim Index As Short = chk.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("CheckBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_Click")
	End Sub
	
	Private Sub chk_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chk.Enter
		Dim Index As Short = chk.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("CheckBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_GotFocus")
	End Sub
	
	Private Sub chk_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles chk.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Dim Index As Short = chk.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("CheckBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_KeyDown", KeyCode, Shift)
	End Sub
	
	Private Sub chk_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles chk.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = chk.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("CheckBox", Index)
		If (RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_KeyPress", KeyAscii)) Then
			' vetoed
			KeyAscii = 0
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub chk_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles chk.KeyUp
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Dim Index As Short = chk.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("CheckBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_KeyUp", KeyCode, Shift)
	End Sub
	
	Private Sub chk_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chk.Leave
		Dim Index As Short = chk.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("CheckBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_LostFocus")
	End Sub
	
	Private Sub chk_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles chk.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim Index As Short = chk.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("CheckBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_MouseDown", Button, Shift, x, y)
	End Sub
	
	Private Sub chk_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles chk.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim Index As Short = chk.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("CheckBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_MouseMode", Button, Shift, x, y)
	End Sub
	
	Private Sub chk_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles chk.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim Index As Short = chk.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("CheckBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_MouseUp", Button, Shift, x, y)
	End Sub
	
	'UPGRADE_WARNING: Event cmb.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	'UPGRADE_WARNING: ComboBox event cmb.Change was upgraded to cmb.TextChanged which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
	Private Sub cmb_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmb.TextChanged
		Dim Index As Short = cmb.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("ComboBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_Change")
	End Sub
	
	'UPGRADE_WARNING: Event cmb.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cmb_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmb.SelectedIndexChanged
		Dim Index As Short = cmb.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("ComboBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_Click")
	End Sub
	
	'UPGRADE_ISSUE: ComboBox event cmb.DblClick was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub cmb_DblClick(ByRef Index As Short)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("ComboBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_DblClick")
	End Sub
	
	Private Sub cmb_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmb.Enter
		Dim Index As Short = cmb.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("ComboBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_GotFocus")
	End Sub
	
	Private Sub cmb_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles cmb.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Dim Index As Short = cmb.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("ComboBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_KeyDown", KeyCode, Shift)
	End Sub
	
	Private Sub cmb_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles cmb.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = cmb.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("ComboBox", Index)
		If (RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_KeyPress", KeyAscii)) Then
			' vetoed
			KeyAscii = 0
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub cmb_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles cmb.KeyUp
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Dim Index As Short = cmb.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("ComboBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_KeyUp", KeyCode, Shift)
	End Sub
	
	Private Sub cmb_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmb.Leave
		Dim Index As Short = cmb.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("ComboBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_LostFocus")
	End Sub
	
	'UPGRADE_ISSUE: ComboBox event cmb.Scroll was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub cmb_Scroll(ByRef Index As Short)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("ComboBox", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_Scroll")
	End Sub
	
	Private Sub trv_AfterLabelEdit(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.NodeLabelEditEventArgs) Handles trv.AfterLabelEdit
		Dim Cancel As Boolean = eventArgs.CancelEdit
		Dim NewString As String = eventArgs.Label
		Dim Index As Short = trv.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("TreeView", Index)
		If (RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_AfterLabelEdit", NewString)) Then
			' vetoed
			Cancel = 1
		End If
	End Sub
	
	Private Sub trv_BeforeLabelEdit(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.NodeLabelEditEventArgs) Handles trv.BeforeLabelEdit
		Dim Cancel As Boolean = eventArgs.CancelEdit
		Dim Index As Short = trv.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("TreeView", Index)
		If (RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_BeforeLabelEdit")) Then
			' vetoed
			Cancel = 1
		End If
	End Sub
	
	Private Sub trv_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles trv.Click
		Dim Index As Short = trv.GetIndex(eventSender)
		
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("TreeView", Index)
		
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_Click")
	End Sub
	
	Private Sub trv_AfterCollapse(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.TreeViewEventArgs) Handles trv.AfterCollapse
		Dim node As System.Windows.Forms.TreeNode = eventArgs.Node
		Dim Index As Short = trv.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("TreeView", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_Collapse", node)
		
	End Sub
	
	Private Sub trv_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles trv.DoubleClick
		Dim Index As Short = trv.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("TreeView", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_DblClick")
	End Sub
	
	Private Sub trv_AfterExpand(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.TreeViewEventArgs) Handles trv.AfterExpand
		Dim node As System.Windows.Forms.TreeNode = eventArgs.Node
		Dim Index As Short = trv.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("TreeView", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_Expand", node)
		
	End Sub
	
	Private Sub trv_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles trv.Enter
		Dim Index As Short = trv.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("TreeView", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_GotFocus")
	End Sub
	
	Private Sub trv_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles trv.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Dim Index As Short = trv.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("TreeView", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_KeyDown", KeyCode, Shift)
	End Sub
	
	Private Sub trv_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles trv.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = trv.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("TreeView", Index)
		If (RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_KeyPress", KeyAscii)) Then
			' vetoed
			KeyAscii = 0
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub trv_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles trv.KeyUp
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Dim Index As Short = trv.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("TreeView", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_KeyUp", KeyCode, Shift)
	End Sub
	
	Private Sub trv_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles trv.Leave
		Dim Index As Short = trv.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("TreeView", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_LostFocus")
	End Sub
	
	Private Sub trv_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles trv.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim Index As Short = trv.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("TreeView", Index)
		
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_MouseDown", Button, Shift, x, y)
	End Sub
	
	Private Sub trv_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles trv.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim Index As Short = trv.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("TreeView", Index)
		
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_MouseMove", Button, Shift, x, y)
	End Sub
	
	Private Sub trv_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles trv.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim Index As Short = trv.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("TreeView", Index)
		
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_MouseUp", Button, Shift, x, y)
	End Sub
	
	Private Sub trv_AfterCheck(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.TreeViewEventArgs) Handles trv.AfterCheck
		Dim node As System.Windows.Forms.TreeNode = eventArgs.Node
		Dim Index As Short = trv.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("TreeView", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_NodeCheck", node)
		
	End Sub
	
	Private Sub trv_NodeClick(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.TreeNodeMouseClickEventArgs) Handles trv.NodeMouseClick
		Dim node As System.Windows.Forms.TreeNode = eventArgs.Node
		Dim Index As Short = trv.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("TreeView", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_NodeClick", node)
		
	End Sub
	
	Private Sub prg_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles prg.Click
		Dim Index As Short = prg.GetIndex(eventSender)
		
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("ProgressBar", Index)
		
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_Click")
	End Sub
	
	Private Sub prg_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles prg.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim Index As Short = prg.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("ProgressBar", Index)
		
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_MouseDown", Button, Shift, x, y)
	End Sub
	
	Private Sub prg_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles prg.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim Index As Short = prg.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("ProgressBar", Index)
		
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_MouseMove", Button, Shift, x, y)
	End Sub
	
	Private Sub prg_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles prg.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim Index As Short = prg.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("ProgressBar", Index)
		
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_MouseUp", Button, Shift, x, y)
	End Sub
	
	'UPGRADE_ISSUE: Frame event fra.Click was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub fra_Click(ByRef Index As Short)
		
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("Frame", Index)
		
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_Click")
	End Sub
	
	'UPGRADE_ISSUE: Frame event fra.DblClick was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub fra_DblClick(ByRef Index As Short)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("Frame", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_DblClick")
	End Sub
	
	'UPGRADE_ISSUE: Frame event fra.MouseDown was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub fra_MouseDown(ByRef Index As Short, ByRef Button As Short, ByRef Shift As Short, ByRef x As Single, ByRef y As Single)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("Frame", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_MouseDown", Button, Shift, x, y)
	End Sub
	
	'UPGRADE_ISSUE: Frame event fra.MouseMove was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub fra_MouseMove(ByRef Index As Short, ByRef Button As Short, ByRef Shift As Short, ByRef x As Single, ByRef y As Single)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("Frame", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_MouseMove", Button, Shift, x, y)
	End Sub
	
	'UPGRADE_ISSUE: Frame event fra.MouseUp was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub fra_MouseUp(ByRef Index As Short, ByRef Button As Short, ByRef Shift As Short, ByRef x As Single, ByRef y As Single)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("Frame", Index)
		RunInSingle(m_sc_module, m_name & "_" & obj.ObjName & "_MouseUp", Button, Shift, x, y)
	End Sub
End Class