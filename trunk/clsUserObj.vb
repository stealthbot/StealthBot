Option Strict Off
Option Explicit On
Friend Class clsUserObj
	' clsUserObj.cls
	' Copyright (C) 2008 Eric Evans
	
	
	
	Private m_flags As Integer
	Private m_ping As Integer
	Private m_actual_name As String
	Private m_character_name As String
	Private m_clan As String
	Private m_clan_rank As Short
	Private m_join_date As Date
	Private m_last_speak_date As Date
	Private m_stats_string As String
	Private m_Game As String
	Private m_total_bans As Integer
	Private m_total_kicks As Integer
	Private m_operator_date As Date
	Private m_queue As Collection
	Private m_passed_chan_auth As Boolean
	Private m_stats As clsUserStats
	Private m_pending_ban As Boolean
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		m_queue = New Collection
		m_stats = New clsUserStats
		
		LastTalkTime = UtcNow
		JoinTime = UtcNow
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		ClearQueue()
		
		'UPGRADE_NOTE: Object m_queue may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_queue = Nothing
		'UPGRADE_NOTE: Object m_stats may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_stats = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	
	Public Property Name() As String
		Get
			Name = modEvents.CleanUsername(m_actual_name)
		End Get
		Set(ByVal Value As String)
			m_actual_name = Value
		End Set
	End Property
	
	
	Public Property CharacterName() As String
		Get
			CharacterName = m_character_name
		End Get
		Set(ByVal Value As String)
			m_character_name = Value
		End Set
	End Property
	
	Public ReadOnly Property Game() As String
		Get
			Game = m_stats.Game
		End Get
	End Property
	
	
	Public Property PendingBan() As Boolean
		Get
			PendingBan = m_pending_ban
		End Get
		Set(ByVal Value As Boolean)
			m_pending_ban = Value
		End Set
	End Property
	
	Public ReadOnly Property IsUsingDII() As Boolean
		Get
			IsUsingDII = ((Game = PRODUCT_D2DV) Or (Game = PRODUCT_D2XP))
		End Get
	End Property
	
	Public ReadOnly Property IsUsingWarIII() As Boolean
		Get
			IsUsingWarIII = ((Game = PRODUCT_WAR3) Or (Game = PRODUCT_W3XP))
		End Get
	End Property
	
	
	Public Property Statstring() As String
		Get
			Statstring = m_stats_string
		End Get
		Set(ByVal Value As String)
			m_stats_string = Value
			
			m_stats.Statstring = m_stats_string
		End Set
	End Property
	
	Public ReadOnly Property Clan() As String
		Get
			Clan = m_stats.Clan
		End Get
	End Property
	
	Public ReadOnly Property DisplayName() As String
		Get
			DisplayName = ConvertUsername(m_actual_name)
		End Get
	End Property
	
	' Converts the username to always contain the gateway
	'  3 = gateway convention: show all
	Public ReadOnly Property FullName() As String
		Get
			FullName = ConvertUsername(m_actual_name, 3)
		End Get
	End Property
	
	
	Public Property Flags() As Integer
		Get
			Flags = m_flags
		End Get
		Set(ByVal Value As Integer)
			m_flags = Value
		End Set
	End Property
	
	
	Public Property PassedChannelAuth() As Boolean
		Get
			PassedChannelAuth = m_passed_chan_auth
		End Get
		Set(ByVal Value As Boolean)
			m_passed_chan_auth = Value
		End Set
	End Property
	
	Public ReadOnly Property IsBnetAdmin() As Boolean
		Get
			IsBnetAdmin = ((m_flags And USER_SYSOP) = USER_SYSOP)
		End Get
	End Property
	
	Public ReadOnly Property IsBlizzRep() As Boolean
		Get
			IsBlizzRep = ((m_flags And USER_BLIZZREP) = USER_BLIZZREP)
		End Get
	End Property
	
	Public ReadOnly Property IsOperator() As Boolean
		Get
			IsOperator = (((m_flags And USER_CHANNELOP) = USER_CHANNELOP) Or IsBlizzRep() Or IsBnetAdmin())
		End Get
	End Property
	
	Public ReadOnly Property IsSquelched() As Boolean
		Get
			IsSquelched = ((m_flags And USER_SQUELCHED) = USER_SQUELCHED)
		End Get
	End Property
	
	
	Public Property Ping() As Integer
		Get
			Ping = m_ping
		End Get
		Set(ByVal Value As Integer)
			m_ping = Value
		End Set
	End Property
	
	
	Public Property LastTalkTime() As Date
		Get
			LastTalkTime = m_last_speak_date
		End Get
		Set(ByVal Value As Date)
			m_last_speak_date = Value
		End Set
	End Property
	
	
	Public Property JoinTime() As Date
		Get
			JoinTime = m_join_date
		End Get
		Set(ByVal Value As Date)
			m_join_date = Value
		End Set
	End Property
	
	Public ReadOnly Property Queue() As Collection
		Get
			Queue = m_queue
		End Get
	End Property
	
	Public ReadOnly Property Stats() As clsUserStats
		Get
			Stats = m_stats
		End Get
	End Property
	
	Public Function TimeSinceTalk() As Double
		On Error GoTo ERROR_HANDLER
		
		'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
		TimeSinceTalk = DateDiff(Microsoft.VisualBasic.DateInterval.Second, LastTalkTime, UtcNow)
		
		Exit Function
		
ERROR_HANDLER: 
		Exit Function
	End Function
	
	Public Function TimeInChannel() As Double
		On Error GoTo ERROR_HANDLER
		
		'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
		TimeInChannel = DateDiff(Microsoft.VisualBasic.DateInterval.Second, JoinTime, UtcNow)
		
		Exit Function
		
ERROR_HANDLER: 
		Exit Function
	End Function
	
	Public Sub ClearQueue()
		Dim i As Short
		
		For i = Queue.Count() To 1 Step -1
			Queue.Remove(i)
		Next i
	End Sub
	
	Public Sub DisplayQueue()
		On Error GoTo ERROR_HANDLER
		
		Dim CurrentEvent As clsUserEventObj
		Dim j As Short
		
		If (Queue Is Nothing) Then
			Exit Sub
		End If
		
		For j = 1 To Queue.Count()
			If (j > Queue.Count()) Then
				Exit For
			End If
			
			CurrentEvent = Queue.Item(j)
			
			Select Case (CurrentEvent.EventID)
				Case ID_USER
					Call Event_UserInChannel(Name, CurrentEvent.Flags, CurrentEvent.Statstring, CurrentEvent.Ping, j)
					
				Case ID_JOIN
					Call Event_UserJoins(Name, CurrentEvent.Flags, CurrentEvent.Statstring, CurrentEvent.Ping, j)
					
				Case ID_TALK
					Call Event_UserTalk(Name, CurrentEvent.Flags, CurrentEvent.Message, CurrentEvent.Ping, j)
					
				Case ID_EMOTE
					Call Event_UserEmote(Name, CurrentEvent.Flags, CurrentEvent.Message, j)
					
				Case ID_USERFLAGS
					Call Event_FlagsUpdate(Name, CurrentEvent.Flags, CurrentEvent.Statstring, CurrentEvent.Ping, j)
			End Select
		Next j
		
		ClearQueue()
		
		Exit Sub
		
ERROR_HANDLER: 
		frmChat.AddChat(RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & " in clsUserObj::DisplayQueue().")
		
		Exit Sub
	End Sub
	
	Public Function Clone() As Object
		Dim i As Short
		
		Clone = New clsUserObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Clone.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Clone.Name = Name
		'UPGRADE_WARNING: Couldn't resolve default property of object Clone.Ping. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Clone.Ping = Ping
		'UPGRADE_WARNING: Couldn't resolve default property of object Clone.Flags. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Clone.Flags = Flags
		'UPGRADE_WARNING: Couldn't resolve default property of object Clone.CharacterName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Clone.CharacterName = CharacterName
		'UPGRADE_WARNING: Couldn't resolve default property of object Clone.JoinTime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Clone.JoinTime = JoinTime
		'UPGRADE_WARNING: Couldn't resolve default property of object Clone.LastTalkTime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Clone.LastTalkTime = LastTalkTime
		'UPGRADE_WARNING: Couldn't resolve default property of object Clone.PassedChannelAuth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Clone.PassedChannelAuth = PassedChannelAuth
		'UPGRADE_WARNING: Couldn't resolve default property of object Clone.PendingBan. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Clone.PendingBan = PendingBan
		'UPGRADE_WARNING: Couldn't resolve default property of object Clone.Statstring. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Clone.Statstring = Statstring
		
		For i = 1 To Queue.Count()
			'UPGRADE_WARNING: Couldn't resolve default property of object Queue().Clone. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Clone.Queue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Clone.Queue.Add(Queue.Item(i).Clone())
		Next i
	End Function
End Class