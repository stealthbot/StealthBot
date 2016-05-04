Option Strict Off
Option Explicit On
Friend Class clsChannelObj
	' clsUserObj.cls
	' Copyright (C) 2008 Eric Evans
	
	
	Private Const CHAN_PUBLIC As Integer = &H1
	Private Const CHAN_MODERATED As Integer = &H2
	Private Const CHAN_RESTRICTED As Integer = &H4
	Private Const CHAN_SILENT As Integer = &H8
	Private Const CHAN_SYSTEM As Integer = &H10
	Private Const CHAN_PRODUCT As Integer = &H20
	Private Const CHAN_GLOBAL As Integer = &H1000
	
	Private m_name As String
	Private m_flags As Integer
	Private m_designated_heir As String
	Private m_num_joins As Integer
	Private m_num_bans As Integer
	Private m_num_kicks As Integer
	Private m_join_date As Date
	Private m_users As Collection
	Private m_banlist As Collection
	
	Public ReadOnly Property SType() As String
		Get
			
			Dim tmp As String
			
			If ((Flags And CHAN_RESTRICTED) = CHAN_RESTRICTED) Then
				tmp = tmp & "restricted, "
			End If
			
			If ((Flags And CHAN_GLOBAL) = CHAN_GLOBAL) Then
				tmp = tmp & "global, "
			End If
			
			If ((Flags And CHAN_PUBLIC) = CHAN_PUBLIC) Then
				tmp = tmp & "public, "
			End If
			
			If ((Flags And CHAN_MODERATED) = CHAN_MODERATED) Then
				tmp = tmp & "moderated, "
			End If
			
			If ((Flags And CHAN_PRODUCT) = CHAN_PRODUCT) Then
				tmp = tmp & "product-specific, "
			End If
			
			If ((Flags And CHAN_SYSTEM) = CHAN_SYSTEM) Then
				tmp = tmp & "system, "
			End If
			
			If ((Flags And CHAN_SILENT) = CHAN_SILENT) Then
				tmp = tmp & "silent, "
			End If
			
			If (Flags = &H0) Then
				tmp = "private, "
			End If
			
			tmp = Mid(tmp, 1, Len(tmp) - 2)
			
			SType = tmp
			
		End Get
	End Property
	
	
	Public Property Name() As String
		Get
			
			Name = m_name
			
		End Get
		Set(ByVal Value As String)
			
			m_name = Value
			
		End Set
	End Property
	
	
	Public Property OperatorHeir() As String
		Get
			
			OperatorHeir = m_designated_heir
			
		End Get
		Set(ByVal Value As String)
			
			m_designated_heir = Value
			
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
	
	
	Public Property JoinTime() As Date
		Get
			
			JoinTime = m_join_date
			
		End Get
		Set(ByVal Value As Date)
			
			m_join_date = Value
			
		End Set
	End Property
	
	Public ReadOnly Property IsSilent() As Boolean
		Get
			
			IsSilent = ((m_flags And CHAN_SILENT) = CHAN_SILENT)
			
		End Get
	End Property
	
	
	Public Property JoinCount() As Integer
		Get
			
			
		End Get
		Set(ByVal Value As Integer)
			
			
		End Set
	End Property
	
	
	Public Property BanCount() As Integer
		Get
			
			BanCount = m_num_bans
			
		End Get
		Set(ByVal Value As Integer)
			
			m_num_bans = Value
			
		End Set
	End Property
	
	
	Public Property KickCount() As Integer
		Get
			
			
		End Get
		Set(ByVal Value As Integer)
			
			
		End Set
	End Property
	
	Public ReadOnly Property Users() As Collection
		Get
			
			Users = m_users
			
		End Get
	End Property
	
	Public ReadOnly Property Self() As clsUserObj
		Get
			
			Dim i As Short
			
			Self = New clsUserObj
			
			For i = 1 To Users.Count()
				'UPGRADE_WARNING: Couldn't resolve default property of object Users().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (StrComp(Users.Item(i).Name, CleanUsername(CurrentUsername), CompareMethod.Text) = 0) Then
					Self = Users.Item(i)
					
					Exit Property
				End If
			Next i
			
		End Get
	End Property
	
	Public ReadOnly Property Banlist() As Collection
		Get
			
			Banlist = m_banlist
			
		End Get
	End Property
	
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		
		m_users = New Collection
		m_banlist = New Collection
		
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		
		'UPGRADE_NOTE: Object m_users may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_users = Nothing
		'UPGRADE_NOTE: Object m_banlist may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_banlist = Nothing
		
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	Public Sub ClearUsers()
		
		'UPGRADE_NOTE: Object m_users may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_users = Nothing
		
		m_users = New Collection
		
	End Sub
	
	Public Sub ClearBanlist()
		
		m_banlist = New Collection
		
	End Sub
	
	Public Function GetUserEx(ByVal AccountName As String, Optional ByVal SearchLimit As Short = 0) As Object
		
		Dim Index As Short
		
		If (StrictIsNumeric(AccountName)) Then
			Index = CShort(Val(AccountName))
		Else
			Index = GetUserIndexEx(AccountName, SearchLimit)
		End If
		
		If ((Index >= 1) And (Index <= Users.Count())) Then
			GetUserEx = Users.Item(Index)
		Else
			GetUserEx = New clsUserObj
		End If
		
	End Function
	
	Public Function GetUserIndexEx(ByVal AccountName As String, Optional ByVal SearchLimit As Short = 0) As Short
		
		Dim i As Short
		
		AccountName = modEvents.CleanUsername(AccountName)
		
		For i = 1 To Users.Count()
			If (SearchLimit > 0) Then
				If (i >= SearchLimit) Then
					Exit For
				End If
			End If
			
			'UPGRADE_WARNING: Couldn't resolve default property of object Users().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If (StrComp(Users.Item(i).Name, AccountName, CompareMethod.Text) = 0) Then
				GetUserIndexEx = i
				
				Exit Function
			End If
		Next i
		
		GetUserIndexEx = 0
		
	End Function
	
	Public Function GetUser(ByVal Username As String, Optional ByVal SearchLimit As Short = 0) As Object
		
		Dim Index As Short
		
		If (StrictIsNumeric(Username) And (Len(Username) <= 4)) Then
			Index = CShort(Val(Username))
			If (Index > Users.Count()) Then
				Index = GetUserIndex(Username, SearchLimit)
			End If
		Else
			Index = GetUserIndex(Username, SearchLimit)
		End If
		
		If ((Index >= 1) And (Index <= Users.Count())) Then
			GetUser = Users.Item(Index)
		Else
			GetUser = New clsUserObj
		End If
		
	End Function
	
	Public Function GetUserIndex(ByVal Username As String, Optional ByVal SearchLimit As Short = 0) As Short
		
		Dim i As Short
		
		For i = 1 To m_users.Count()
			If (SearchLimit > 0) Then
				If (i >= SearchLimit) Then
					Exit For
				End If
			End If
			
			'UPGRADE_WARNING: Couldn't resolve default property of object m_users().DisplayName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If (StrComp(m_users.Item(i).DisplayName, Username, CompareMethod.Text) = 0) Then
				GetUserIndex = i
				
				Exit Function
			End If
		Next i
		
		GetUserIndex = 0
		
	End Function
	
	'UPGRADE_NOTE: Operator was upgraded to Operator_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function IsOnBanList(ByVal Username As String, Optional ByVal Operator_Renamed As String = vbNullString) As Short
		
		Dim i As Short
		Dim bln As Boolean
		
		Username = CleanUsername(Username)
		
		For i = 1 To Banlist.Count()
			'UPGRADE_WARNING: Couldn't resolve default property of object Banlist().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If (StrComp(Banlist.Item(i).Name, Username, CompareMethod.Text) = 0) Then
				If (Operator_Renamed <> vbNullString) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Banlist().Operator. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If (StrComp(Banlist.Item(i).Operator, Operator_Renamed, CompareMethod.Text) = 0) Then
						bln = True
					End If
				Else
					bln = True
				End If
				
				If (bln) Then
					IsOnBanList = i
					
					Exit Function
				End If
			End If
		Next i
		
		IsOnBanList = 0
		
	End Function
	
	'UPGRADE_NOTE: Operator was upgraded to Operator_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function IsOnRecentBanList(ByVal Username As String, Optional ByVal Operator_Renamed As String = vbNullString) As Short
		
		Dim i As Short
		Dim bln As Boolean
		
		Username = CleanUsername(Username)
		
		For i = Banlist.Count() To (Banlist.Count() - 5) Step -1
			If (i <= 0) Then
				Exit For
			End If
			
			'UPGRADE_WARNING: Couldn't resolve default property of object Banlist().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If (StrComp(Banlist.Item(i).Name, Username, CompareMethod.Text) = 0) Then
				If (Operator_Renamed <> vbNullString) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Banlist().Operator. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If (StrComp(Banlist.Item(i).Operator, Operator_Renamed, CompareMethod.Text) = 0) Then
						bln = True
					End If
				Else
					bln = True
				End If
				
				If (bln) Then
					'frmChat.AddChat vbYellow, DateDiff("s", Banlist(I).DateOfBan, UtcNow)
					
					'UPGRADE_WARNING: Couldn't resolve default property of object Banlist().DateOfBan. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
					If (DateDiff(Microsoft.VisualBasic.DateInterval.Second, Banlist.Item(i).DateOfBan, UtcNow) <= 3) Then
						IsOnRecentBanList = i
						
						Exit Function
					End If
				End If
			End If
		Next i
		
		IsOnRecentBanList = 0
		
	End Function
	
	Public Sub RemoveBansFromOperator(ByVal Username As String)
		
		Dim i As Short
		Dim bln As Boolean
		Dim pos As Short
		Dim Name As String
		
		Username = CleanUsername(Username)
		
		Do 
			bln = False
			
			For i = 1 To Banlist.Count()
				'UPGRADE_WARNING: Couldn't resolve default property of object Banlist().Operator. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (StrComp(Banlist.Item(i).Operator, Username, CompareMethod.Text) = 0) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Banlist().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Name = Banlist.Item(i).Name
					
					If (BotVars.RetainOldBans = False) Then
						Banlist.Remove(i)
						
						pos = IsOnBanList(Name)
						
						If (pos > 0) Then
							'UPGRADE_WARNING: Couldn't resolve default property of object Banlist().IsDuplicateBan. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							Banlist.Item(pos).IsDuplicateBan = False
						End If
						
						bln = True
						
						Exit For
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object Banlist().IsActive. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Banlist.Item(i).IsActive = False
					End If
				End If
			Next i
		Loop While (bln = True)
		
	End Sub
	
	Public Function CheckUser(ByRef Username As String, Optional ByRef CurrentUser As clsUserObj = Nothing) As Short
		
		Dim doCheck As Boolean
		
		doCheck = True
		
		Dim DBEntry As udtGetAccessResponse
		Dim i As Short
		Dim Message As String
		Dim j As Short
		If (Self.IsOperator) Then
			
			If (CurrentUser Is Nothing) Then
				CurrentUser = Users.Item(GetUserIndex(Username))
			End If
			
			'UPGRADE_WARNING: Couldn't resolve default property of object DBEntry. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			DBEntry = GetCumulativeAccess(CurrentUser.DisplayName, "USER")
			
			If (DBEntry.Rank < AutoModSafelistValue) Then
				' Not high enough rank to be safe from auto-mod
				
				If (InStr(1, DBEntry.Flags, "S", CompareMethod.Binary) = 0) Then
					' Not on the safelist
					
					If (InStr(1, DBEntry.Flags, "B", CompareMethod.Binary) <> 0) Then
						' On the shitlist
						
						Message = GetShitlist(CurrentUser.DisplayName)
						
						frmChat.AddQ("/ban " & Message)
						
						doCheck = False
					Else
						
						' Is channel protection enabled?
						If (Protect) Then
							Ban(CurrentUser.DisplayName & Space(1) & ProtectMsg, AutoModSafelistValue - 1)
							
							doCheck = False
						Else
							
							' Have they previously been banned?
							If ((doCheck) And (BotVars.BanEvasion)) Then
								If (IsOnBanList(CurrentUser.Name)) Then
									Ban(CurrentUser.DisplayName & " Ban Evasion", AutoModSafelistValue - 1)
									
									doCheck = False
								End If
							End If
							
							' Are they IP banned?
							If ((doCheck) And (BotVars.IPBans)) Then
								If (CurrentUser.IsSquelched) Then
									Ban(CurrentUser.DisplayName & " IP Banned", AutoModSafelistValue - 1)
									
									doCheck = False
								End If
							End If
							
							' Is the user's ping acceptable?
							If ((doCheck) And (Config.PingBan)) Then
								If (Config.PingBanLevel < 1) Then
									If (CurrentUser.Ping = Config.PingBanLevel) Then
										Ban(CurrentUser.DisplayName & " Ping Ban", AutoModSafelistValue - 1)
										
										doCheck = False
									End If
								Else
									If (CurrentUser.Ping > Config.PingBanLevel) Then
										Ban(CurrentUser.DisplayName & " Ping Ban", AutoModSafelistValue - 1)
										
										doCheck = False
									End If
								End If
							End If
							
							' Do they have a UDP plug?
							If ((doCheck) And (BotVars.PlugBan)) Then
								If ((CurrentUser.Flags And USER_NOUDP) = USER_NOUDP) Then
									Ban(CurrentUser.DisplayName & " PlugBan", AutoModSafelistValue - 1)
									
									doCheck = False
								End If
							End If
							
							' Level bans
							If (doCheck) Then
								If (CurrentUser.IsUsingDII) Then
									If (BotVars.BanD2UnderLevel) Then
										If (CurrentUser.Stats.Level < BotVars.BanD2UnderLevel) Then
											Message = BotVars.BanUnderLevelMsg
											
                                            If Len(Message) = 0 Then
                                                Message = "You are below the required level for entry."
                                            End If
											
											If InStr(1, Message, "%cl", CompareMethod.Text) > 0 Then
												Message = Replace(Message, "%cl", CStr(CurrentUser.Stats.Level))
											End If
											
											If InStr(1, Message, "%rl", CompareMethod.Text) > 0 Then
												Message = Replace(Message, "%rl", CStr(BotVars.BanD2UnderLevel))
											End If
											
											Ban(CurrentUser.DisplayName & Space(1) & Message, AutoModSafelistValue - 1)
											
											doCheck = False
										End If
									End If
								ElseIf (CurrentUser.IsUsingWarIII) Then 
									If (BotVars.BanPeons) Then
										If (StrComp(CurrentUser.Stats.IconName, "peon", CompareMethod.Text) = 0) Then
											Ban(CurrentUser.DisplayName & Space(1) & "PeonBan", AutoModSafelistValue - 1)
											
											doCheck = False
										End If
									ElseIf (BotVars.BanUnderLevel) Then 
										If (CurrentUser.Stats.Level < BotVars.BanUnderLevel) Then
											Message = BotVars.BanUnderLevelMsg
											
                                            If Len(Message) = 0 Then
                                                Message = "You are below the required level for entry."
                                            End If
											
											If InStr(1, Message, "%cl", CompareMethod.Text) > 0 Then
												Message = Replace(Message, "%cl", CStr(CurrentUser.Stats.Level))
											End If
											
											If InStr(1, Message, "%rl", CompareMethod.Text) > 0 Then
												Message = Replace(Message, "%rl", CStr(BotVars.BanUnderLevel))
											End If
											
											Ban(CurrentUser.DisplayName & Space(1) & Message, AutoModSafelistValue - 1)
											
											doCheck = False
										End If
									End If
								End If
							End If
							
							'If (doCheck) Then
							'    If ((BotVars.ChannelPasswordDelay) And (Len(BotVars.ChannelPassword) > 0)) Then
							'        If (CurrentUser.TimeInChannel() > BotVars.ChannelPasswordDelay) Then
							'            Ban CurrentUser.DisplayName & " Password time is up", _
							''                (AutoModSafelistValue - 1)
							'
							'            doCheck = False
							'        End If
							'    End If
							'End If
							'
							'If (doCheck) Then
							'    If ((BotVars.IB_On = BTRUE) And (BotVars.IB_Wait > 0)) Then
							'        If (CurrentUser.TimeSinceTalk() > BotVars.IB_Wait) Then
							'            Ban CurrentUser.DisplayName & " Idle for " & BotVars.IB_Wait & "+ seconds", _
							''                (AutoModSafelistValue - 1), IIf(BotVars.IB_Kick, 1, 0)
							'
							'            doCheck = False
							'        End If
							'    End If
							'End If
						End If
					End If
				End If
			End If
		End If
		
		If (doCheck = False) Then
			CurrentUser.PendingBan = True
		End If
		
	End Function
	
	Public Function CheckUsers() As Short
		
		Dim CurrentUser As clsUserObj
		Dim DBEntry As udtGetAccessResponse
		Dim i As Short
		Dim Message As String
		Dim doCheck As Boolean
		Dim HighRank As Short
		Dim HighIndex As Short
		Dim Count As Short
		If (Self.IsOperator) Then
			
			doCheck = True
			
			If (OperatorHeir = vbNullString) Then
				For i = 1 To Users.Count()
					CurrentUser = Users.Item(i)
					
					If (CurrentUser.IsOperator = False) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object DBEntry. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						DBEntry = GetCumulativeAccess(CurrentUser.DisplayName, "USER")
						
						If (InStr(1, DBEntry.Flags, "D", CompareMethod.Binary) <> 0) Then
							If (DBEntry.Rank > HighRank) Then
								HighRank = DBEntry.Rank
								
								HighIndex = i
							End If
						End If
					End If
				Next i
				
				If (HighIndex > 0) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Users().DisplayName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					frmChat.AddQ("/designate " & Users.Item(HighIndex).DisplayName)
				End If
			End If
			
			For i = Users.Count() To 1 Step -1
				CurrentUser = Users.Item(i)
				
				If (CurrentUser.IsOperator = False) Then
					If (CheckUser((CurrentUser.DisplayName), CurrentUser)) Then
						Count = (Count + 1)
					End If
				End If
			Next i
		End If
		
		CheckUsers = Count
		
	End Function
	
	Public Function CheckQueue(ByVal Username As String) As Boolean
		
		Dim CurrentUser As clsUserObj
		
		For	Each CurrentUser In Users
			If (StrComp(CurrentUser.DisplayName, Username, CompareMethod.Text) = 0) Then
				If (CurrentUser.Queue.Count()) Then
					CheckQueue = True
				End If
				
				Exit For
			End If
		Next CurrentUser
		
	End Function
	
	Public Function Clone() As Object
		
		Dim i As Short
		
		Clone = New clsChannelObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Clone.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Clone.Name = Name
		'UPGRADE_WARNING: Couldn't resolve default property of object Clone.Flags. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Clone.Flags = Flags
		'UPGRADE_WARNING: Couldn't resolve default property of object Clone.JoinTime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Clone.JoinTime = JoinTime
		'UPGRADE_WARNING: Couldn't resolve default property of object Clone.BanCount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Clone.BanCount = BanCount
		'UPGRADE_WARNING: Couldn't resolve default property of object Clone.KickCount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Clone.KickCount = KickCount
		'UPGRADE_WARNING: Couldn't resolve default property of object Clone.JoinCount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Clone.JoinCount = JoinCount
		'UPGRADE_WARNING: Couldn't resolve default property of object Clone.OperatorHeir. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Clone.OperatorHeir = OperatorHeir
		
		For i = 1 To Users.Count()
			'UPGRADE_WARNING: Couldn't resolve default property of object Users().Clone. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Clone.Users. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Clone.Users.Add(Users.Item(i).Clone())
		Next i
		
		For i = 1 To Banlist.Count()
			'UPGRADE_WARNING: Couldn't resolve default property of object Banlist().Clone. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Clone.Banlist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Clone.Banlist.Add(Banlist.Item(i).Clone())
		Next i
		
	End Function
End Class