Option Strict Off
Option Explicit On
Friend Class clsFriendObj
	'---------------------------------------------------------------------------------------
	' Module    : g_FriendObj
	' DateTime  : 3/15/2004 18:58
	' Author    : Stealth
	' Purpose   : stores friendlist data
	'---------------------------------------------------------------------------------------
	
	
	Private m_Username As String
	Private m_status As Byte
	Private m_location_id As Byte
	Private m_location As String
	Private m_game As String
	
	
	Public Property Location() As String
		Get
			Location = m_location
		End Get
		Set(ByVal Value As String)
			m_location = Value
		End Set
	End Property
	
	
	Public Property game() As String
		Get
			game = m_game
		End Get
		Set(ByVal Value As String)
			m_game = KillNull(Value)
		End Set
	End Property
	
	
	Public Property LocationID() As Byte
		Get
			LocationID = m_location_id
		End Get
		Set(ByVal Value As Byte)
			m_location_id = Value
		End Set
	End Property
	
	
	Public Property Status() As Byte
		Get
			Status = m_status
		End Get
		Set(ByVal Value As Byte)
			m_status = Value
		End Set
	End Property
	
	Public ReadOnly Property IsMutual() As Boolean
		Get
			IsMutual = (Status And 1)
		End Get
	End Property
	
	Public ReadOnly Property IsAway() As Boolean
		Get
			IsAway = (Status And 4)
		End Get
	End Property
	
	Public ReadOnly Property DontDisturb() As Boolean
		Get
			DontDisturb = (Status And 2)
		End Get
	End Property
	
	Public ReadOnly Property IsOnline() As Boolean
		Get
			IsOnline = (LocationID <> 0)
		End Get
	End Property
	
	Public ReadOnly Property IsInChannel() As Boolean
		Get
			IsInChannel = (LocationID = 2)
		End Get
	End Property
	
	Public ReadOnly Property IsInGame() As Boolean
		Get
			IsInGame = (LocationID > 2)
		End Get
	End Property
	
	
	Public Property Name() As String
		Get
			Name = m_Username
		End Get
		Set(ByVal Value As String)
			m_Username = KillNull(Value)
		End Set
	End Property
	
	Public ReadOnly Property DisplayName() As String
		Get
			DisplayName = ConvertUsername(m_Username)
		End Get
	End Property
	
	Public Function Clone() As Object
		Clone = New clsFriendObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Clone.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Clone.Name = Name
		'UPGRADE_WARNING: Couldn't resolve default property of object Clone.game. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Clone.game = game
		'UPGRADE_WARNING: Couldn't resolve default property of object Clone.Status. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Clone.Status = Status
		'UPGRADE_WARNING: Couldn't resolve default property of object Clone.LocationID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Clone.LocationID = LocationID
		'UPGRADE_WARNING: Couldn't resolve default property of object Clone.Location. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Clone.Location = Location
	End Function
End Class