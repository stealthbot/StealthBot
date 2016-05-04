Option Strict Off
Option Explicit On
Friend Class clsQueueOBj
	
	Private m_obj_id As Double
	Private m_message As String
	Private m_priority As Short
	Private m_response As String
	Private m_tag As String
	
	
	Public Property id() As Double
		Get
			
			id = m_obj_id
			
		End Get
		Set(ByVal Value As Double)
			
			If (m_obj_id > 0) Then
				Exit Property
			End If
			
			m_obj_id = Value
			
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
	
	
	Public Property PRIORITY_Renamed() As Short
		Get
			
			PRIORITY = m_priority
			
		End Get
		Set(ByVal Value As Short)
			
			m_priority = Value
			
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
	
	
	Public Property ResponseTo() As String
		Get
			
			ResponseTo = m_response
			
		End Get
		Set(ByVal Value As String)
			
			m_response = Value
			
		End Set
	End Property
End Class