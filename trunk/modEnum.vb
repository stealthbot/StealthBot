Option Strict Off
Option Explicit On
Module modEnum
	
	'modEnum - project StealthBot
	'February 2004 - Stealth [stealth at stealthbot dot net]
	
	
	'UDTS
	Public Structure udtChanList
		Dim Current As String
		Dim Flags As Integer
		Dim Designated As String
		Dim staticDesignee As String
	End Structure
	
	Public Structure udtOutFilters
		Dim ofFind As String
		Dim ofReplace As String
	End Structure
	
	Public Structure udtBanList
		Dim Username As String
		Dim UsernameActual As String
		Dim cOperator As String
	End Structure
	
	Public Structure udtAutoRespond
		Dim Check As String
		Dim Reply As String
	End Structure
	
	Public Structure udtProductInfo
		Dim Code As String
		Dim ShortCode As String
		Dim KeyCount As Short
		Dim FullName As String
		Dim BNLS_ID As Integer
		Dim LogonSystem As Integer
		Dim VersionByte As Integer
	End Structure
	
	Public Structure udtCustomCommandData
		Dim reqAccess As Short
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(20),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=20)> Public Query() As Char
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(500),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=500)> Public Action() As Char
	End Structure
	
	Public Structure udtDatabase
		Dim Username As String
		Dim Rank As Short
		Dim Flags As String
		Dim AddedBy As String
		Dim AddedOn As Date
		Dim ModifiedBy As String
		Dim ModifiedOn As Date
		Dim Type As String
		Dim Groups As String
		Dim BanMessage As String
	End Structure
	
	'Public Type udtQueue
	'    Message     As String
	'    Priority    As Byte
	'End Type
	
	Public Structure udtGetAccessResponse
		Dim Username As String
		Dim Rank As Short
		Dim Flags As String
		Dim AddedBy As String
		Dim AddedOn As Date
		Dim ModifiedBy As String
		Dim ModifiedOn As Date
		Dim Type As String
		Dim Groups As String
		Dim BanMessage As String
	End Structure
	
	Public Structure udtMail
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		'UPGRADE_NOTE: To was upgraded to To_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		<VBFixedString(30),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=30)> Public To_Renamed() As Char
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(30),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=30)> Public From() As Char
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(225),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=225)> Public Message() As Char
	End Structure
	
	Public Structure FLASHWINFO
		Dim cbSize As Integer
		Dim hWnd As Integer
		Dim dwFlags As Integer
		Dim uCount As Integer
		Dim dwTimeout As Integer
	End Structure
	
	Public Structure POINTAPI
		Dim x As Integer
		Dim y As Integer
	End Structure
	
	Public Structure LVHITTESTINFO
		Dim pt As POINTAPI
		Dim Flags As Integer
		Dim iItem As Integer
		Dim iSubItem As Integer
	End Structure
	
	Public Enum inetQueueModes
		inqReset = 0
		inqAdd = 1
		inqGet = 2
	End Enum
	
	Public Enum enuErrorSources
		BNET = 0
		BNLS = 1
		MCP = 2
	End Enum
	
	Public Enum enuProxyStatus
		psNotConnected = 0
		psConnecting = 1
		psLoggingIn = 2
		psOnline = 3
	End Enum
	
	Public Enum enuWebProfileTypes
		W3XP = 1
		WAR3 = 2
	End Enum
	
	Public Enum eCharacterTypes
		Unknown = &H0
		Amazon = &H1
		Sorceress = &H2
		Necromancer = &H3
		Paladin = &H4
		Barbarian = &H5
		Druid = &H6
		Assassin = &H7
	End Enum
	
	Public Enum enuSettingsPanels
		spConnectionConfig = 0
		spConnectionAdvanced = 1
		spInterfaceGeneral = 2
		spInterfaceFontsColors = 3
		spGenModeration = 4
		spGenGreets = 5
		spGenIdles = 6
		spGenMisc = 7
		spSplash = 8
	End Enum
	
	Public Enum enuDBActions
		AddEntry = 1
		RemEntry = 2
		ModEntry = 3
	End Enum
	
	Public Enum enuPL_ServerTypes
		stBNLS = 1
		stBNCS = 2
		stMCP = 3
	End Enum
	
	Public Enum enuPL_DirectionTypes
		CtoS = 1
		StoC = 2
	End Enum
End Module