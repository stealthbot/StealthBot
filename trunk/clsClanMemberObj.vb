Option Strict Off
Option Explicit On
Friend Class clsClanMemberObj
	' clsClanMemberObj.cls
	' Copyright (C) 2008 Eric Evans
	
	
	Private m_name As String
	Private m_rank As Short
	Private m_join_date As Date
	Private m_status As Short
	Private m_location As String
	
	
	Public Property Name() As String
		Get
			Name = m_name
		End Get
		Set(ByVal Value As String)
			m_name = Value
		End Set
	End Property
	
	Public ReadOnly Property DisplayName() As String
		Get
			DisplayName = ConvertUsername(Name)
		End Get
	End Property
	
	
	Public Property Rank() As Short
		Get
			Rank = m_rank
		End Get
		Set(ByVal Value As Short)
			m_rank = Value
		End Set
	End Property
	
	Public ReadOnly Property RankName() As String
		Get
			RankName = GetRank(CByte(m_rank))
		End Get
	End Property
	
	
	Public Property JoinTime() As Date
		Get
			JoinTime = m_join_date
		End Get
		Set(ByVal Value As Date)
			m_join_date = Value
		End Set
	End Property
	
	
	Public Property Status() As Short
		Get
			Status = m_status
		End Get
		Set(ByVal Value As Short)
			m_status = Value
		End Set
	End Property
	
	Public ReadOnly Property IsOnline() As Boolean
		Get
			IsOnline = (Status > 0)
		End Get
	End Property
	
	
	Public Property Location() As String
		Get
			Location = m_location
		End Get
		Set(ByVal Value As String)
			m_location = Value
		End Set
	End Property
	
	Public Sub MakeChieftain()
		Call MakeMemberChieftain(m_name)
	End Sub
	
	Public Sub Promote(Optional ByVal Rank As Short = -1)
		If ((Rank > -1) And (Rank <= m_rank)) Then
			Exit Sub
		End If
		Call PromoteMember(m_name, IIf(Rank > -1, Rank, m_rank + 1))
	End Sub
	
	Public Sub Demote(Optional ByVal Rank As Short = -1)
		If ((Rank > -1) And (Rank >= m_rank)) Then
			Exit Sub
		End If
		Call DemoteMember(m_name, IIf(Rank > -1, Rank, m_rank - 1))
	End Sub
	
	Public Sub KickOut()
		Call RemoveMember(m_name)
	End Sub
	
	Public Function Clone() As Object
		Clone = New clsClanMemberObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Clone.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Clone.Name = Name
		'UPGRADE_WARNING: Couldn't resolve default property of object Clone.Location. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Clone.Location = Location
		'UPGRADE_WARNING: Couldn't resolve default property of object Clone.Rank. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Clone.Rank = Rank
		'UPGRADE_WARNING: Couldn't resolve default property of object Clone.Status. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Clone.Status = Status
		'UPGRADE_WARNING: Couldn't resolve default property of object Clone.JoinTime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Clone.JoinTime = JoinTime
	End Function
End Class