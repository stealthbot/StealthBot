Option Strict Off
Option Explicit On
Friend Class clsClanObj
	' clsMyClanObj.cls
	' Copyright (C) 2008 Eric Evans
	
	
	Private m_name As String
	Private m_motd As String
	Private m_members As Collection
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		m_members = New Collection
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'UPGRADE_NOTE: Object m_members may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_members = Nothing
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
	
	
	Public Property MOTD() As String
		Get
			MOTD = m_motd
		End Get
		Set(ByVal Value As String)
			m_motd = Value
		End Set
	End Property
	
	Public ReadOnly Property Self() As clsClanMemberObj
		Get
			Dim i As Short
			
			Self = New clsClanMemberObj
			
			For i = 1 To Members.Count()
				'Changed 10/08/09 - Hdx - Now uses BotVars.Username instead of CurrentUsername so that if you login as a multiple (#2 #2) it will still work.
				'UPGRADE_WARNING: Couldn't resolve default property of object Members().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (StrComp(Members.Item(i).Name, BotVars.Username, CompareMethod.Text) = 0) Then
					Self = Members.Item(i)
					Exit For
				End If
			Next i
		End Get
	End Property
	
	Public ReadOnly Property Members() As Collection
		Get
			Members = m_members
		End Get
	End Property
	
	Public ReadOnly Property Chieftain() As clsClanMemberObj
		Get
			Dim i As Short
			Chieftain = New clsClanMemberObj
			
			For i = 1 To Members.Count()
				'UPGRADE_WARNING: Couldn't resolve default property of object Members(i).Rank. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (Members.Item(i).Rank >= 4) Then
					Chieftain = Members.Item(i)
					
					Exit For
				End If
			Next i
		End Get
	End Property
	
	Public ReadOnly Property Shamans() As Collection
		Get
			Dim i As Short
			Shamans = New Collection
			
			For i = 1 To Members.Count()
				'UPGRADE_WARNING: Couldn't resolve default property of object Members(i).Rank. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (Members.Item(i).Rank = 3) Then
					Shamans.Add(Members.Item(i))
				End If
			Next i
		End Get
	End Property
	
	Public ReadOnly Property Grunts() As Collection
		Get
			Dim i As Short
			Grunts = New Collection
			
			For i = 1 To Members.Count()
				'UPGRADE_WARNING: Couldn't resolve default property of object Members(i).Rank. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (Members.Item(i).Rank = 2) Then
					Grunts.Add(Members.Item(i))
				End If
			Next i
		End Get
	End Property
	
	Public ReadOnly Property Peons() As Collection
		Get
			Dim i As Short
			Peons = New Collection
			
			For i = 1 To Members.Count()
				'UPGRADE_WARNING: Couldn't resolve default property of object Members(i).Rank. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If ((Members.Item(i).Rank = 0) Or (Members.Item(i).Rank = 1)) Then
					Peons.Add(Members.Item(i))
				End If
			Next i
		End Get
	End Property
	
	Public Function GetMember(ByVal Username As String) As Object
		GetMember = GetUser(Username)
	End Function
	
	Public Function GetUser(ByVal Username As String) As Object
		Dim Index As Short
		
		If (StrictIsNumeric(Username)) Then
			Index = CShort(Val(Username))
		Else
			Index = GetUserIndex(Username)
		End If
		
		If ((Index >= 1) And (Index <= Members.Count())) Then
			GetUser = Members.Item(Index)
		Else
			GetUser = New clsClanMemberObj
		End If
	End Function
	
	Public Function GetMemberIndex(ByVal Username As String) As Short
		GetMemberIndex = GetUserIndex(Username)
	End Function
	
	Public Function GetUserIndex(ByVal Username As String) As Short
		Dim i As Short
		
		For i = 1 To Members.Count()
			'UPGRADE_WARNING: Couldn't resolve default property of object Members().DisplayName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If (StrComp(Members.Item(i).DisplayName, Username, CompareMethod.Text) = 0) Then
				GetUserIndex = i
				
				Exit Function
			End If
		Next i
		
		GetUserIndex = 0
	End Function
	
	Public Function GetMemberEx(ByVal Username As String) As Object
		GetMemberEx = GetUserEx(Username)
	End Function
	
	Public Function GetUserEx(ByVal Username As String) As Object
		Dim Index As Short
		
		If (StrictIsNumeric(Username)) Then
			Index = CShort(Val(Username))
		Else
			Index = GetUserIndexEx(Username)
		End If
		
		If ((Index >= 1) And (Index <= Members.Count())) Then
			GetUserEx = Members.Item(Index)
		Else
			GetUserEx = New clsClanMemberObj
		End If
	End Function
	
	Public Function GetMemberIndexEx(ByVal Username As String) As Short
		GetMemberIndexEx = GetUserIndexEx(Username)
	End Function
	
	Public Function GetUserIndexEx(ByVal Username As String) As Short
		Dim i As Short
		
		For i = 1 To Members.Count()
			'UPGRADE_WARNING: Couldn't resolve default property of object Members().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If (StrComp(Members.Item(i).Name, Username, CompareMethod.Text) = 0) Then
				GetUserIndexEx = i
				
				Exit Function
			End If
		Next i
		
		GetUserIndexEx = 0
	End Function
	
	Public Sub Clear()
		m_members = New Collection
	End Sub
	
	Public Sub Disband()
		Call DisbandClan()
	End Sub
	
	Public Sub SetMOTD(ByVal MOTD As String)
		Call modWar3Clan.SetClanMOTD(MOTD)
	End Sub
	
	Public Function Clone() As Object
		Dim i As Short
		Clone = New clsClanObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Clone.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Clone.Name = Name
		'UPGRADE_WARNING: Couldn't resolve default property of object Clone.MOTD. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Clone.MOTD = MOTD
		
		For i = 1 To Members.Count()
			'UPGRADE_WARNING: Couldn't resolve default property of object Members().Clone. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Clone.Members. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Clone.Members.Add(Members.Item(i).Clone())
		Next i
	End Function
End Class