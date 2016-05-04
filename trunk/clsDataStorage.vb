Option Strict Off
Option Explicit On
Friend Class clsDataStorage
	Private m_lNLSHandle As Integer
	
	Private m_ServerToken As Integer
	Private m_ClientToken As Integer
	Private m_lLogonType As Integer
	Private m_UDPValue As Integer
	Private m_CRevFileTime As String
	Private m_CRevFileName As String
	Private m_CRevSeed As String
	Private m_CRevVersion As Integer
	Private m_CRevChecksum As Integer
	Private m_CRevResult As String
	Private m_ServerSig As String
	Private m_EmailRegDelay As Boolean
	Private m_NLS As clsNLS
	Private m_MCPHandler As clsMCPHandler
	Private m_FirstTimeChat As Boolean
	
	Public Sub List()
		With frmChat
			.AddChat(RTBColors.ErrorMessageText, StringFormat("Logon Type:   0x{0}", ZeroOffset(LogonType, 8)))
			.AddChat(RTBColors.ErrorMessageText, StringFormat("Server Token: 0x{0}", ZeroOffset(ServerToken, 8)))
			.AddChat(RTBColors.ErrorMessageText, StringFormat("Client Token: 0x{0}", ZeroOffset(ClientToken, 8)))
			.AddChat(RTBColors.ErrorMessageText, StringFormat("UDP Value:    0x{0}", ZeroOffset(UDPValue, 8)))
			.AddChat(RTBColors.ErrorMessageText, "CRev Info: ")
			.AddChat(RTBColors.ErrorMessageText, StringFormat("  FileTime: {0}", CRevFileTime))
			.AddChat(RTBColors.ErrorMessageText, StringFormat("  FileName: {0}", CRevFileName))
			.AddChat(RTBColors.ErrorMessageText, StringFormat("  Seed:     {0}", IIf(InStr(1, CRevSeed, "A=", CompareMethod.Text) = 0, StrToHex(CRevSeed), CRevSeed)))
			.AddChat(RTBColors.ErrorMessageText, StringFormat("  Version:  0x{0}", ZeroOffset(CRevVersion, 8)))
			.AddChat(RTBColors.ErrorMessageText, StringFormat("  Checksum: 0x{0}", ZeroOffset(CRevChecksum, 8)))
			.AddChat(RTBColors.ErrorMessageText, StringFormat("  Result:   {0}", IIf(InStr(1, CRevSeed, "A=", CompareMethod.Text) = 0, StrToHex(CRevResult), CRevResult)))
			'.AddChat RTBColors.ErrorMessageText, StringFormat("MCP Data:{0}{1}", vbNewLine, DebugOutput(m_MCPData))
		End With
	End Sub
	
	'UPGRADE_NOTE: Reset was upgraded to Reset_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub Reset_Renamed()
		m_ServerToken = 0
		m_ClientToken = 0
		m_lLogonType = 0
		m_UDPValue = 0
		m_CRevFileTime = vbNullString
		m_CRevFileName = vbNullString
		m_CRevSeed = vbNullString
		m_CRevVersion = 0
		m_CRevChecksum = 0
		m_CRevResult = vbNullString
		m_ServerSig = vbNullString
		m_EmailRegDelay = False
		'UPGRADE_NOTE: Object m_MCPHandler may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_MCPHandler = Nothing
		'UPGRADE_NOTE: Object m_NLS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_NLS = Nothing
		m_FirstTimeChat = False
	End Sub
	
	Public Property LogonType() As Integer
		Get
			LogonType = m_lLogonType
		End Get
		Set(ByVal Value As Integer)
			m_lLogonType = Value
		End Set
	End Property
	
	Public Property ServerToken() As Integer
		Get
			ServerToken = m_ServerToken
		End Get
		Set(ByVal Value As Integer)
			m_ServerToken = Value
		End Set
	End Property
	
	Public Property ClientToken() As Integer
		Get
			If (m_ClientToken = 0) Then m_ClientToken = GetTickCount
			ClientToken = m_ClientToken
		End Get
		Set(ByVal Value As Integer)
			m_ClientToken = Value
		End Set
	End Property
	
	Public Property UDPValue() As Integer
		Get
			UDPValue = m_ServerToken
		End Get
		Set(ByVal Value As Integer)
			m_UDPValue = Value
		End Set
	End Property
	
	Public Property CRevFileTime() As String
		Get
			Dim ft As FILETIME
			Dim st As SYSTEMTIME
			
			If (Not Len(m_CRevFileTime) = 8) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				m_CRevFileTime = Left(StringFormat("{0}{1}", m_CRevFileTime, New String(Chr(0), 8)), 8)
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object ft. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(ft, m_CRevFileTime, 8)
			
			FileTimeToSystemTime(ft, st)
			With st
				'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				CRevFileTime = StringFormat("{0}/{1}/{2} {3}:{4}:{5}", .wMonth, .wDay, .wYear, .wHour, .wMinute, .wSecond)
			End With
		End Get
		Set(ByVal Value As String)
			m_CRevFileTime = Value
		End Set
	End Property
	Public ReadOnly Property CRevFileTimeRaw() As String
		Get
			CRevFileTimeRaw = m_CRevFileTime
		End Get
	End Property
	
	Public Property CRevFileName() As String
		Get
			CRevFileName = m_CRevFileName
		End Get
		Set(ByVal Value As String)
			m_CRevFileName = Value
		End Set
	End Property
	
	Public Property CRevSeed() As String
		Get
			CRevSeed = m_CRevSeed
		End Get
		Set(ByVal Value As String)
			m_CRevSeed = Value
		End Set
	End Property
	
	Public Property CRevResult() As String
		Get
			CRevResult = m_CRevResult
		End Get
		Set(ByVal Value As String)
			m_CRevResult = Value
		End Set
	End Property
	
	Public Property CRevVersion() As Integer
		Get
			CRevVersion = m_CRevVersion
		End Get
		Set(ByVal Value As Integer)
			m_CRevVersion = Value
		End Set
	End Property
	
	Public Property CRevChecksum() As Integer
		Get
			CRevChecksum = m_CRevChecksum
		End Get
		Set(ByVal Value As Integer)
			m_CRevChecksum = Value
		End Set
	End Property
	
	Public Property ServerSig() As String
		Get
			ServerSig = m_ServerSig
		End Get
		Set(ByVal Value As String)
			m_ServerSig = Value
		End Set
	End Property
	
	Public Property WaitingForEmail() As Boolean
		Get
			WaitingForEmail = m_EmailRegDelay
		End Get
		Set(ByVal Value As Boolean)
			m_EmailRegDelay = Value
		End Set
	End Property
	
	Public ReadOnly Property NLS() As clsNLS
		Get
			If (m_NLS Is Nothing) Then m_NLS = New clsNLS
			NLS = m_NLS
		End Get
	End Property
	
	Public Property MCPHandler() As clsMCPHandler
		Get
			MCPHandler = m_MCPHandler
		End Get
		Set(ByVal Value As clsMCPHandler)
			m_MCPHandler = Value
		End Set
	End Property
	
	Public Property EnteredChatFirstTime() As Boolean
		Get
			EnteredChatFirstTime = m_FirstTimeChat
		End Get
		Set(ByVal Value As Boolean)
			m_FirstTimeChat = Value
		End Set
	End Property
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		Reset_Renamed()
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
End Class