Option Strict Off
Option Explicit On
Friend Class clsCommandRestrictionObj
	' clsCommandRestrictionObj.cls
	' Copyright (C) 2008 Eric Evans
	
	
	Private m_name As String
	Private m_req_rank As Short
	Private m_req_flags As String
	Private m_matchmessage As String
	Private m_casesensitive As Boolean
	Private m_error As String
	Private m_fatal As Boolean
	
	
	Public Property Name() As String
		Get
			Name = m_name
		End Get
		Set(ByVal Value As String)
			m_name = Value
		End Set
	End Property
	
	
	Public Property RequiredRank() As Short
		Get
			RequiredRank = m_req_rank
		End Get
		Set(ByVal Value As Short)
			m_req_rank = Value
		End Set
	End Property
	
	
	Public Property RequiredFlags() As String
		Get
			RequiredFlags = m_req_flags
		End Get
		Set(ByVal Value As String)
			m_req_flags = Value
		End Set
	End Property
	
	
	Public Property MatchMessage() As String
		Get
			MatchMessage = m_matchmessage
		End Get
		Set(ByVal Value As String)
			m_matchmessage = Value
		End Set
	End Property
	
	
	Public Property MatchError() As String
		Get
			MatchError = m_error
		End Get
		Set(ByVal Value As String)
			m_error = Value
		End Set
	End Property
	
	
	Public Property MatchCaseSensitive() As Boolean
		Get
			MatchCaseSensitive = m_casesensitive
		End Get
		Set(ByVal Value As Boolean)
			m_casesensitive = Value
		End Set
	End Property
	
	
	Public Property Fatal() As Boolean
		Get
			Fatal = m_fatal
		End Get
		Set(ByVal Value As Boolean)
			m_fatal = Value
		End Set
	End Property
End Class