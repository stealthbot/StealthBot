Option Strict Off
Option Explicit On
Friend Class clsDBEntryObj
	' clsDBEntryObj.cls
	' Copyright (C) 2008 Eric Evans
	
	
	Private m_type As String
	Private m_name As String
	Private m_rank As Integer
	Private m_flags As String
	Private m_created_on As Date
	Private m_created_by As String
	Private m_modified_on As Date
	Private m_modified_by As String
	Private m_groups As Collection
	Private m_lastseen As Date
	Private m_ban_message As String
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		m_groups = New Collection
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	
	Public Property EntryType() As String
		Get
			EntryType = m_type
		End Get
		Set(ByVal Value As String)
			m_type = Value
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
	
	
	Public Property Rank() As Integer
		Get
			Rank = m_rank
		End Get
		Set(ByVal Value As Integer)
			m_rank = Value
		End Set
	End Property
	
	
	Public Property Flags() As String
		Get
			Flags = m_flags
		End Get
		Set(ByVal Value As String)
			m_flags = Value
		End Set
	End Property
	
	
	Public Property CreatedOn() As Date
		Get
			CreatedOn = m_created_on
		End Get
		Set(ByVal Value As Date)
			m_created_on = Value
		End Set
	End Property
	
	
	Public Property CreatedBy() As String
		Get
			CreatedBy = m_created_by
		End Get
		Set(ByVal Value As String)
			m_created_by = Value
		End Set
	End Property
	
	
	Public Property ModifiedOn() As Date
		Get
			ModifiedOn = m_modified_on
		End Get
		Set(ByVal Value As Date)
			m_modified_on = Value
		End Set
	End Property
	
	
	Public Property ModifiedBy() As String
		Get
			ModifiedBy = m_modified_by
		End Get
		Set(ByVal Value As String)
			m_modified_by = Value
		End Set
	End Property
	
	
	Public Property LastSeen() As Date
		Get
			LastSeen = m_lastseen
		End Get
		Set(ByVal Value As Date)
			m_lastseen = Value
		End Set
	End Property
	
	Public ReadOnly Property Groups() As Collection
		Get
			Groups = m_groups
		End Get
	End Property
	
	'What exactly is this function suposto do?
	'To-Do: Complete function
	Public ReadOnly Property MembersOf() As Collection
		Get
			MembersOf = New Collection
			
			If (StrComp(m_type, "Group", CompareMethod.Text) = 0) Then
				'Do something..
			End If
		End Get
	End Property
	
	
	Public Property BanMessage() As String
		Get
			BanMessage = m_ban_message
		End Get
		Set(ByVal Value As String)
			m_ban_message = Value
		End Set
	End Property
	
	Public Function IsInGroup(ByVal GroupName As String) As Boolean
		Dim i As Short
		Dim pos As Short
		
		For i = 1 To Groups.Count()
			'UPGRADE_WARNING: Couldn't resolve default property of object Groups().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pos = InStr(1, Groups.Item(i).Name, Space(1), CompareMethod.Binary)
			
			If (pos > 0) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Groups(i).Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object Groups().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Groups.Item(i).Name = Mid(Groups.Item(i).Name, 1, pos - 1)
			End If
			
			'UPGRADE_WARNING: Couldn't resolve default property of object Groups().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If (StrComp(GroupName, Groups.Item(i).Name, CompareMethod.Text) = 0) Then
				IsInGroup = True
				
				Exit Function
			End If
		Next i
	End Function
	
	Public Function HasFlag(ByVal strFlag As String, Optional ByVal CaseSensitive As Boolean = True) As Boolean
		If (CaseSensitive) Then
			HasFlag = (InStr(1, m_flags, strFlag, CompareMethod.Binary) <> 0)
		Else
			HasFlag = (InStr(1, m_flags, strFlag, CompareMethod.Text) <> 0)
		End If
	End Function
	
	Public Function HasAnyFlag(ByVal strFlags As String, Optional ByVal CaseSensitive As Boolean = True) As Boolean
		Dim i As Short
		
		For i = 1 To Len(strFlags)
			If (CaseSensitive) Then
				HasAnyFlag = (InStr(1, m_flags, Mid(strFlags, i, 1), CompareMethod.Binary) <> 0)
			Else
				HasAnyFlag = (InStr(1, m_flags, Mid(strFlags, i, 1), CompareMethod.Text) <> 0)
			End If
			
			If (HasAnyFlag) Then
				Exit Function
			End If
		Next i
	End Function
	
	Public Function HasFlags(ByVal strFlags As String, Optional ByVal CaseSensitive As Boolean = True) As Boolean
		Dim i As Short
		
		For i = 1 To Len(strFlags)
			If (CaseSensitive) Then
				HasFlags = (InStr(1, m_flags, Mid(strFlags, i, 1), CompareMethod.Binary) <> 0)
			Else
				HasFlags = (InStr(1, m_flags, Mid(strFlags, i, 1), CompareMethod.Text) <> 0)
			End If
			
			If (HasFlags = False) Then
				Exit Function
			End If
		Next i
	End Function
	
	Public Sub AddGroup(ByVal sGroup As String)
		m_groups.Add(sGroup, sGroup)
	End Sub
End Class