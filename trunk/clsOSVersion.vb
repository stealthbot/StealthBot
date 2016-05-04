Option Strict Off
Option Explicit On
Friend Class clsOSVersion
	'modOSVersion.bas
	' project StealthBot
	' October 2006 from code at:
	'  http://vbnet.mvps.org/index.html?code/helpers/iswinversion.htm
	
	
    Private Declare Function GetVersionEx Lib "Kernel32.dll" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Integer
    Private Declare Function GetVersionEx Lib "Kernel32.dll" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFOEX) As Integer
	
	Private Const VER_PLATFORM_WIN32_NT As Integer = 2
	Private Const VER_NT_WORKSTATION As Integer = 1
	Private Const VER_NT_DOMAIN_CONTROLLER As Integer = 2
	Private Const VER_NT_SERVER As Integer = 3
	
	Private Structure OSVERSIONINFO
		Dim OSVSize As Integer 'size, in bytes, of this data structure
		Dim dwVerMajor As Integer 'ie NT 3.51, dwVerMajor = 3; NT 4.0, dwVerMajor = 4.
		Dim dwVerMinor As Integer 'ie NT 3.51, dwVerMinor = 51; NT 4.0, dwVerMinor= 0.
		Dim dwBuildNumber As Integer 'NT: build number of the OS
		'Win9x: build number of the OS in low-order word.
		'       High-order word contains major & minor ver nos.
		Dim PlatformID As Integer 'Identifies the operating system platform.
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(128),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=128)> Public szCSDVersion() As Char 'NT: string, such as "Service Pack 3"
		'Win9x: string providing arbitrary additional information
	End Structure
	
	Private Structure OSVERSIONINFOEX
		Dim OSVSize As Integer
		Dim dwVerMajor As Integer
		Dim dwVerMinor As Integer
		Dim dwBuildNumber As Integer
		Dim PlatformID As Integer
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(128),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=128)> Public szCSDVersion() As Char
		Dim wServicePackMajor As Short
		Dim wServicePackMinor As Short
		Dim wSuiteMask As Short
		Dim wProductType As Byte
		Dim wReserved As Byte
	End Structure
	
	Private m_isCached As Boolean
	Private m_osVer As OSVERSIONINFO
	Private m_osVerEx As OSVERSIONINFOEX
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		GetVersion()
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		m_isCached = False
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	Private Sub GetVersion()
		Dim bln As Boolean
		
		m_osVer.OSVSize = Len(m_osVer)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object m_osVer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (GetVersionEx(m_osVer) = 1) Then
			If (IsWindowsNT) Then
				m_osVerEx.OSVSize = Len(m_osVerEx)
				
				'UPGRADE_WARNING: Couldn't resolve default property of object m_osVerEx. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				m_isCached = (GetVersionEx(m_osVerEx) = 1)
				
				Exit Sub
			End If
			
			m_isCached = True
		End If
	End Sub
	
	Public ReadOnly Property Name() As String
		Get
			If (IsWindowsNT = False) Then
				If (IsWindows95) Then
					Name = "Windows 95"
				ElseIf (IsWindows98) Then 
					Name = "Windows 98"
				ElseIf (IsWindowsME) Then 
					Name = "Windows ME"
				End If
			Else
				If (IsWindows2000) Then
					Name = "Windows 2000"
				ElseIf (IsWindowsXP) Then 
					Name = "Windows XP"
				ElseIf (IsWindows2003) Then 
					Name = "Windows Server 2003"
				ElseIf (IsWindowsVista) Then 
					Name = "Windows Vista"
				ElseIf (IsWindows2008) Then 
					Name = "Windows Server 2008"
				ElseIf (IsWindows7) Then 
					Name = "Windows 7"
				Else
					Name = "Windows NT " & m_osVerEx.dwVerMajor & "." & m_osVerEx.dwVerMinor
				End If
			End If
			
			If (Name = vbNullString) Then
				Name = "Unknown"
			End If
		End Get
	End Property
	
	Public ReadOnly Property IsWindowsNT() As Boolean
		Get
			IsWindowsNT = CBool(m_osVer.PlatformID = VER_PLATFORM_WIN32_NT)
		End Get
	End Property
	
	Public ReadOnly Property IsWindows95() As Boolean
		Get
			IsWindows95 = CBool((m_osVer.dwVerMajor = 4) And (m_osVer.dwVerMinor = 0))
		End Get
	End Property
	
	Public ReadOnly Property IsWindows98() As Boolean
		Get
			If ((m_osVer.dwVerMajor = 4) And (m_osVer.dwVerMinor = 10)) Then
				IsWindows98 = CBool(m_osVer.dwBuildNumber < 2222)
			End If
		End Get
	End Property
	
	Public ReadOnly Property IsWindowsME() As Boolean
		Get
			If ((m_osVer.dwVerMajor = 4) And (m_osVer.dwVerMinor = 10)) Then
				IsWindowsME = CBool(m_osVer.dwBuildNumber >= 2222)
			End If
		End Get
	End Property
	
	Public ReadOnly Property IsWindows2000() As Boolean
		Get
			IsWindows2000 = CBool((m_osVerEx.dwVerMajor = 5) And (m_osVerEx.dwVerMinor = 0))
		End Get
	End Property
	
	
	Public ReadOnly Property IsWindowsXP() As Boolean
		Get
			IsWindowsXP = CBool((m_osVerEx.dwVerMajor = 5) And (m_osVerEx.dwVerMinor = 1))
			
			' check for winxp 64-bit
			If (IsWindowsXP = False) Then
				If (m_osVerEx.wProductType = VER_NT_WORKSTATION) Then
					IsWindowsXP = CBool((m_osVerEx.dwVerMajor = 5) And (m_osVerEx.dwVerMinor = 2))
				End If
			End If
		End Get
	End Property
	
	Public ReadOnly Property IsWindows2003() As Boolean
		Get
			If (IsWindowsXP = False) Then
				IsWindows2003 = ((m_osVerEx.dwVerMajor = 5) And (m_osVerEx.dwVerMinor = 2))
			End If
		End Get
	End Property
	
	'Added by FrOzeN on 18th September, 2008.
	'Returns true if Vista, false if not.
	' updated by eric
	Public ReadOnly Property IsWindowsVista() As Boolean
		Get
			If (m_osVerEx.wProductType = VER_NT_WORKSTATION) Then
				IsWindowsVista = (CBool(m_osVerEx.dwVerMajor = 6) And (m_osVerEx.dwVerMinor = 0))
			End If
		End Get
	End Property
	
	Public ReadOnly Property IsWindows2008() As Boolean
		Get
			If (m_osVerEx.wProductType <> VER_NT_WORKSTATION) Then
				IsWindows2008 = CBool((m_osVerEx.dwVerMajor = 6) And ((m_osVerEx.dwVerMinor = 0) Or (m_osVerEx.dwVerMinor = 1)))
			End If
		End Get
	End Property
	
	Public ReadOnly Property IsWindows7() As Boolean
		Get
			If (m_osVerEx.wProductType = VER_NT_WORKSTATION) Then
				IsWindows7 = CBool((m_osVerEx.dwVerMajor = 6) And (m_osVerEx.dwVerMinor = 1))
			End If
		End Get
	End Property
	
	Public ReadOnly Property IsWin2000Plus() As Boolean
		Get
			IsWin2000Plus = CBool(m_osVerEx.dwVerMajor >= 5)
		End Get
	End Property
	
	Public ReadOnly Property IsWinVistaPlus() As Boolean
		Get
			IsWinVistaPlus = CBool(m_osVerEx.dwVerMajor >= 6)
		End Get
	End Property
End Class