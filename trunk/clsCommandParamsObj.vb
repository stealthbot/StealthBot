Option Strict Off
Option Explicit On
Friend Class clsCommandParamsObj
	' clsCommandParamsObj.cls
	' Copyright (C) 2008 Eric Evans
	
	
	Private m_name As String
	Private m_optional As Boolean
	Private m_required_rank As Short
	Private m_required_flags As String
	Private m_description As String
	Private m_special_notes As String
	Private m_restrictions As Collection
	Private m_data_type As String
	Private m_matchmessage As String
	Private m_casesensitive As Boolean
	Private m_error As String
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		m_restrictions = New Collection
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'UPGRADE_NOTE: Object m_restrictions may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_restrictions = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	
	Public Function GetRestrictionByName(ByVal sRestrictionName As String) As clsCommandRestrictionObj
		Dim r As clsCommandRestrictionObj
		Dim col As Collection
		Dim i As Short
		
		col = Me.Restrictions
		
		For i = 1 To col.Count()
			r = col.Item(i)
			If StrComp(sRestrictionName, r.Name, CompareMethod.Text) = 0 Then
				GetRestrictionByName = r
				Exit Function
			End If
		Next i
	End Function
	
	
	Public Property Restrictions() As Collection
		Get
			Restrictions = m_restrictions
		End Get
		Set(ByVal Value As Collection)
			m_restrictions = Value
		End Set
	End Property
	
	
	Public Property datatype() As String
		Get
			datatype = m_data_type
		End Get
		Set(ByVal Value As String)
			Select Case LCase(Value)
				Case "string"
				Case "number", "numeric"
				Case "word"
				Case Else
					'// default to string
					Value = "string"
			End Select
			m_data_type = Value
		End Set
	End Property
	
	
	Public Property Name() As String
		Get
			Name = m_name
		End Get
		Set(ByVal Value As String)
			m_name = Value
		End Set
	End Property
	
	
	Public Property description() As String
		Get
			description = m_description
		End Get
		Set(ByVal Value As String)
			m_description = Value
		End Set
	End Property
	
	
	Public Property SpecialNotes() As String
		Get
			SpecialNotes = m_special_notes
		End Get
		Set(ByVal Value As String)
			m_special_notes = Value
		End Set
	End Property
	
	
	Public Property IsOptional() As Boolean
		Get
			IsOptional = m_optional
		End Get
		Set(ByVal Value As Boolean)
			m_optional = Value
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
	
	'Public Property Get Pattern() As String
	'    Pattern = m_data_pattern
	'End Property
	'Public Property Let Pattern(strPattern As String)
	'    m_data_pattern = strPattern
	'End Property
	'Public Property Get min() As Long
	'    min = m_data_min
	'End Property
	'Public Property Let min(Val As Long)
	'    m_data_min = Val
	'End Property
	'Public Property Get Max() As Long
	'    Max = m_data_max
	'End Property
	'Public Property Let Max(Val As Long)
	'    m_data_max = Val
	'End Property
End Class