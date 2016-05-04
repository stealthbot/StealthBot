Option Strict Off
Option Explicit On
Friend Class clsSLongTimer
	
	Private m_tmrObj As Object
	Private m_interval As Short
	Private m_counter As Double
	
	
	Public Property tmr() As Object
		Get
			
			tmr = m_tmrObj
			
		End Get
		Set(ByVal Value As Object)
			
			m_tmrObj = Value
			
		End Set
	End Property
	
	Public ReadOnly Property Parent() As System.Windows.Forms.Form
		Get
			
			'UPGRADE_WARNING: Couldn't resolve default property of object m_tmrObj.Parent. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Parent = m_tmrObj.Parent
			
		End Get
	End Property
	
	Public ReadOnly Property Index() As Short
		Get
			
			'UPGRADE_WARNING: Couldn't resolve default property of object m_tmrObj.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Index = m_tmrObj.Index
			
		End Get
	End Property
	
	Public ReadOnly Property Name() As String
		Get
			
			'UPGRADE_WARNING: Couldn't resolve default property of object m_tmrObj.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Name = m_tmrObj.Name
			
		End Get
	End Property
	
	
	Public Property Tag() As String
		Get
			
			'UPGRADE_WARNING: Couldn't resolve default property of object m_tmrObj.Tag. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Tag = m_tmrObj.Tag
			
		End Get
		Set(ByVal Value As String)
			
			'UPGRADE_WARNING: Couldn't resolve default property of object m_tmrObj.Tag. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_tmrObj.Tag = Value
			
		End Set
	End Property
	
	
	Public Property Interval() As Short
		Get
			
			Interval = m_interval
			
		End Get
		Set(ByVal Value As Short)
			
			m_interval = Value
			
		End Set
	End Property
	
	
	Public Property Enabled() As Boolean
		Get
			
			'UPGRADE_WARNING: Couldn't resolve default property of object m_tmrObj.Enabled. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Enabled = m_tmrObj.Enabled
			
		End Get
		Set(ByVal Value As Boolean)
			
			'UPGRADE_WARNING: Couldn't resolve default property of object m_tmrObj.Enabled. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_tmrObj.Enabled = Value
			
		End Set
	End Property
	
	
	Public Property Counter() As Double
		Get
			
			Counter = m_counter
			
		End Get
		Set(ByVal Value As Double)
			
			m_counter = Value
			
		End Set
	End Property
End Class