Option Strict Off
Option Explicit On
Friend Class clsMenuObj
	
	Private m_hWnd As Integer
	Private m_command_id As Integer
	Private m_name As String
	Private m_parent As Object
	Private m_enabled As Boolean
	Private m_visible As Boolean
	Private m_help_id As Integer
	Private m_checked As Boolean
	Private m_tag As String
	Private m_window_list As Boolean
	Private m_subs As Collection
	Private m_has_children As Boolean
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		m_enabled = True
		m_visible = True
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub Class_Terminate_Renamed()
		DeleteMenu()
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	
	Public Property Name() As String
		Get
			Name = m_name
		End Get
		Set(ByVal Value As String)
			m_name = Value
		End Set
	End Property
	
	
	Public Property hWnd() As Integer
		Get
			hWnd = m_hWnd
		End Get
		Set(ByVal Value As Integer)
			If (m_hWnd > 0) Then
				DeleteMenu()
			End If
			
			m_hWnd = Value
			
			CreateMenu()
		End Set
	End Property
	
	
	Public Property HasChildren() As Boolean
		Get
			HasChildren = m_has_children
		End Get
		Set(ByVal Value As Boolean)
			m_has_children = Value
			
			Enabled = Enabled
		End Set
	End Property
	
	
	Public Property Parent() As Object
		Get
			Parent = m_parent
		End Get
		Set(ByVal Value As Object)
			If (Not (m_parent Is Nothing)) Then
				DeleteMenu()
			End If
			
			m_parent = Value
			
			If (TypeOf Value Is System.Windows.Forms.Form) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object obj.hWnd. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				m_hWnd = GetMenu(Value.hWnd)
			ElseIf (TypeOf Value Is clsMenuObj) Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object obj.ID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				m_hWnd = Value.ID
				
				'UPGRADE_WARNING: Couldn't resolve default property of object obj.HasChildren. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Value.HasChildren = True
			End If
			
			CreateMenu()
		End Set
	End Property
	
	
	Public Property ID() As Integer
		Get
			ID = m_command_id
		End Get
		Set(ByVal Value As Integer)
			If (m_command_id = 0) Then
				m_command_id = Value
			End If
		End Set
	End Property
	
	
	Public Property Caption() As String
		Get
			'UPGRADE_NOTE: str was upgraded to str_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
			Dim str_Renamed As String
			Dim lng As Integer
			
			str_Renamed = New String(Chr(0), 256)
			
			lng = GetMenuString(m_hWnd, m_command_id, str_Renamed, Len(str_Renamed), MF_BYCOMMAND)
			
			Caption = Left(str_Renamed, lng)
		End Get
		Set(ByVal Value As String)
			If (Value = "-") Then
				ModifyMenu(m_hWnd, m_command_id, MF_BYCOMMAND Or MF_SEPARATOR, m_command_id, Value)
			Else
				ModifyMenu(m_hWnd, m_command_id, MF_BYCOMMAND, m_command_id, Value)
			End If
			
			DrawMenuBar(m_hWnd)
		End Set
	End Property
	
	
	Public Property Enabled() As Boolean
		Get
			Enabled = m_enabled
		End Get
		Set(ByVal Value As Boolean)
			If (Value = True) Then
				If (ModifyMenu(m_hWnd, m_command_id, MF_BYCOMMAND Or IIf(HasChildren, MF_POPUP, MF_STRING), m_command_id, Caption()) <> -1) Then
					
					m_enabled = True
				End If
			Else
				If (ModifyMenu(m_hWnd, m_command_id, MF_BYCOMMAND Or MF_GRAYED, m_command_id, Caption()) <> -1) Then
					
					m_enabled = False
				End If
			End If
			
			DrawMenuBar(m_hWnd)
		End Set
	End Property
	
	
	Public Property Checked() As Boolean
		Get
			Checked = m_checked
		End Get
		Set(ByVal Value As Boolean)
			If (Value = True) Then
				If (CheckMenuItem(m_hWnd, m_command_id, MF_BYCOMMAND Or MF_CHECKED) <> -1) Then
					m_checked = True
				End If
			Else
				If (CheckMenuItem(m_hWnd, m_command_id, MF_BYCOMMAND Or MF_UNCHECKED) <> -1) Then
					m_checked = False
				End If
			End If
			
			DrawMenuBar(m_hWnd)
		End Set
	End Property
	
	
	Public Property Visible() As Boolean
		Get
			Visible = m_visible
		End Get
		Set(ByVal Value As Boolean)
			If (Value = True) Then
				CreateMenu()
			Else
				DeleteMenu()
			End If
		End Set
	End Property
	
	'To-Do: Complete this Function
	
	'To-Do: Complete this Function
	Public Property HelpContextID() As Integer
		Get
		End Get
		Set(ByVal Value As Integer)
		End Set
	End Property
	
	
	Public Property Tag() As String
		Get
			Tag = m_tag
		End Get
		Set(ByVal Value As String)
			m_tag = Value
		End Set
	End Property
	
	'To-Do: Complete this Function
	Public ReadOnly Property WindowList() As Boolean
		Get
		End Get
	End Property
	
	Private Function CreateMenu() As Integer
		If (m_command_id = 0) Then
			
			m_command_id = CreatePopupMenu()
			
			AppendMenu(m_hWnd, IIf(HasChildren, MF_POPUP, MF_STRING), m_command_id, "")
			
			DrawMenuBar(m_hWnd)
			
			CreateMenu = m_command_id
			
		End If
	End Function
	
	Private Function DeleteMenu() As Integer
		If (m_command_id <> 0) Then
			
			RemoveMenu(m_hWnd, m_command_id, MF_BYCOMMAND)
			
			m_command_id = 0
			
		End If
	End Function
End Class