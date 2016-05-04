Option Strict Off
Option Explicit On
Friend Class clsBanlistUserObj
	' clsBanlistUserObj.cls
	' Copyright (C) 2008 Eric Evans
	
	
	Private m_name As String
	Private m_banned_date As Date
	Private m_ban_reason As String
	Private m_operator As String
	Private m_duplicate_ban As Boolean
	Private m_active_ban As Boolean
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		m_active_ban = True
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	
	Public Property Name() As String
		Get
			Name = modEvents.CleanUsername(m_name)
		End Get
		Set(ByVal Value As String)
			m_name = Value
		End Set
	End Property
	
	Public ReadOnly Property DisplayName() As String
		Get
			DisplayName = convertUsername(Name)
		End Get
	End Property
	
	
	'UPGRADE_NOTE: Operator was upgraded to Operator_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Property Operator_Renamed() As String
		Get
			Operator_Renamed = m_operator
		End Get
		Set(ByVal Value As String)
			m_operator = Value
		End Set
	End Property
	
	
	Public Property DateOfBan() As Date
		Get
			DateOfBan = m_banned_date
		End Get
		Set(ByVal Value As Date)
			m_banned_date = Value
		End Set
	End Property
	
	
	Public Property Reason() As String
		Get
			Reason = m_ban_reason
		End Get
		Set(ByVal Value As String)
			m_ban_reason = Value
		End Set
	End Property
	
	
	Public Property IsDuplicateBan() As Boolean
		Get
			IsDuplicateBan = m_duplicate_ban
		End Get
		Set(ByVal Value As Boolean)
			m_duplicate_ban = Value
		End Set
	End Property
	
	
	Public Property IsActive() As Boolean
		Get
			IsActive = m_active_ban
		End Get
		Set(ByVal Value As Boolean)
			m_active_ban = Value
		End Set
	End Property
	
	Public Function Clone() As Object
		Clone = New clsBanlistUserObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Clone.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Clone.Name = Name
		'UPGRADE_WARNING: Couldn't resolve default property of object Clone.DateOfBan. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Clone.DateOfBan = DateOfBan
		'UPGRADE_WARNING: Couldn't resolve default property of object Clone.IsActive. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Clone.IsActive = IsActive
		'UPGRADE_WARNING: Couldn't resolve default property of object Clone.IsDuplicateBan. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Clone.IsDuplicateBan = IsDuplicateBan
		'UPGRADE_WARNING: Couldn't resolve default property of object Clone.Operator. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Clone.Operator = Operator_Renamed
		'UPGRADE_WARNING: Couldn't resolve default property of object Clone.Reason. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Clone.Reason = Reason
	End Function
End Class