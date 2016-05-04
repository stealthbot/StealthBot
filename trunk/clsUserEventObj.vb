Option Strict Off
Option Explicit On
Friend Class clsUserEventObj
	' clsUserEventObj.cls
	' Copyright (C) 2008 Eric Evans
	
	
	Private m_event_id As Integer
	Private m_gtc As Integer
	Private m_ping As Integer
	Private m_flags As Integer
	Private m_message As String
	Private m_clan As String
	Private m_game_id As String
	Private m_icon_code As String
	Private m_stat_string As String
	Private m_displayed As Boolean
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		EventTick = GetTickCount()
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	
	Public Property EventID() As Integer
		Get
			EventID = m_event_id
		End Get
		Set(ByVal Value As Integer)
			m_event_id = Value
		End Set
	End Property
	
	
	
	Public Property EventTick() As Integer
		Get
			EventTick = m_gtc
		End Get
		Set(ByVal Value As Integer)
			m_gtc = Value
		End Set
	End Property
	
	
	Public Property Ping() As Integer
		Get
			Ping = m_ping
		End Get
		Set(ByVal Value As Integer)
			m_ping = Value
		End Set
	End Property
	
	
	Public Property Flags() As Integer
		Get
			Flags = m_flags
		End Get
		Set(ByVal Value As Integer)
			m_flags = Value
		End Set
	End Property
	
	
	Public Property Message() As String
		Get
			Message = m_message
		End Get
		Set(ByVal Value As String)
			m_message = Value
		End Set
	End Property
	
	
	Public Property GameID() As String
		Get
			GameID = m_game_id
		End Get
		Set(ByVal Value As String)
			m_game_id = Value
		End Set
	End Property
	
	
	Public Property Clan() As String
		Get
			Clan = m_clan
		End Get
		Set(ByVal Value As String)
			m_clan = Value
		End Set
	End Property
	
	
	Public Property Statstring() As String
		Get
			Statstring = m_stat_string
		End Get
		Set(ByVal Value As String)
			m_stat_string = Value
		End Set
	End Property
	
	
	Public Property IconCode() As String
		Get
			IconCode = m_icon_code
		End Get
		Set(ByVal Value As String)
			m_icon_code = Value
		End Set
	End Property
	
	
	Public Property Displayed() As Boolean
		Get
			' returns whether this event has been displayed in the RTB, useful for combining events
			' in messages if chatdelay > 0, such as ops acquired and stats updates
			Displayed = m_displayed
		End Get
		Set(ByVal Value As Boolean)
			' sets whether this event has been displayed in the RTB
			m_displayed = Value
		End Set
	End Property
	
	Public Function Clone() As Object
		Clone = New clsUserEventObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Clone.EventID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Clone.EventID = EventID
		'UPGRADE_WARNING: Couldn't resolve default property of object Clone.EventTick. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Clone.EventTick = EventTick
		'UPGRADE_WARNING: Couldn't resolve default property of object Clone.Flags. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Clone.Flags = Flags
		'UPGRADE_WARNING: Couldn't resolve default property of object Clone.GameID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Clone.GameID = GameID
		'UPGRADE_WARNING: Couldn't resolve default property of object Clone.Clan. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Clone.Clan = Clan
		'UPGRADE_WARNING: Couldn't resolve default property of object Clone.IconCode. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Clone.IconCode = IconCode
		'UPGRADE_WARNING: Couldn't resolve default property of object Clone.Message. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Clone.Message = Message
		'UPGRADE_WARNING: Couldn't resolve default property of object Clone.Ping. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Clone.Ping = Ping
		'UPGRADE_WARNING: Couldn't resolve default property of object Clone.Statstring. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Clone.Statstring = Statstring
	End Function
End Class