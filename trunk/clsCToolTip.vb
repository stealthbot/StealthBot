Option Strict Off
Option Explicit On
Friend Class clsCTooltip
	' This is not my code
	
	Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
	
	''Windows API Functions
    Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Integer, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hWndParent As Integer, ByVal hMenu As Integer, ByVal hInstance As Integer, ByRef lpParam As Integer) As Integer
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As IntPtr, ByVal wMsg As Integer, ByVal wParam As Integer, ByRef lParam As String) As Integer
    Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
    Private Declare Function SendMessageInfo Lib "user32" Alias "SendMessageA" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As TOOLINFO) As Integer
	Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Integer) As Integer
	
	''Windows API Constants
	Private Const WM_USER As Integer = &H400
	Private Const CW_USEDEFAULT As Integer = &H80000000
	
	''Windows API Types
	Private Structure RECT
		'UPGRADE_NOTE: Left was upgraded to Left_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Left_Renamed As Integer
		Dim Top As Integer
		'UPGRADE_NOTE: Right was upgraded to Right_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Right_Renamed As Integer
		Dim Bottom As Integer
	End Structure
	
	''Tooltip Window Constants
	Private Const TTS_NOPREFIX As Integer = &H2
	'Private Const TTF_TRANSPARENT = &H100
	Private Const TTF_CENTERTIP As Integer = &H2
	Private Const TTM_ADDTOOLA As Decimal = (WM_USER + 4)
	'Private Const TTM_ACTIVATE = WM_USER + 1
	Private Const TTM_UPDATETIPTEXTA As Decimal = (WM_USER + 12)
	'Private Const TTM_SETMAXTIPWIDTH = (WM_USER + 24)
	Private Const TTM_SETTIPBKCOLOR As Decimal = (WM_USER + 19)
	Private Const TTM_SETTIPTEXTCOLOR As Decimal = (WM_USER + 20)
	Private Const TTM_SETTITLE As Decimal = (WM_USER + 32)
	Private Const TTS_BALLOON As Integer = &H40
	Private Const TTS_ALWAYSTIP As Integer = &H1
	Private Const TTF_SUBCLASS As Integer = &H10
	Private Const TTF_IDISHWND As Integer = &H1
	Private Const TTM_SETDELAYTIME As Decimal = (WM_USER + 3)
	Private Const TTDT_AUTOPOP As Short = 2
	Private Const TTDT_INITIAL As Short = 3
	
	Private Const TOOLTIPS_CLASSA As String = "tooltips_class32"
	
	''Tooltip Window Types
	Private Structure TOOLINFO
		Dim lSize As Integer
		Dim lFlags As Integer
		Dim hWnd As Integer
		Dim lId As Integer
		Dim lpRect As RECT
		Dim hInstance As Integer
		Dim lpStr As String
		Dim lParam As Integer
	End Structure
	
	
	Public Enum ttIconType
		TTNoIcon = 0
		TTIconInfo = 1
		TTIconWarning = 2
		TTIconError = 3
	End Enum
	
	Public Enum ttStyleEnum
		TTStandard
		TTBalloon
	End Enum
	
	'local variable(s) to hold property value(s)
	Private mvarBackColor As Integer
	Private mvarTitle As String
	Private mvarForeColor As Integer
	Private mvarIcon As ttIconType
	Private mvarCentered As Boolean
	Private mvarStyle As ttStyleEnum
	Private mvarTipText As String
	Private mvarVisibleTime As Integer
	Private mvarDelayTime As Integer
	
	'private data
	Private m_lTTHwnd As Integer ' hwnd of the tooltip
	Private m_lParentHwnd As Integer ' hwnd of the window the tooltip attached to
	Private ti As TOOLINFO
	
	
	Public Property Style() As ttStyleEnum
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Style
			Style = mvarStyle
		End Get
		Set(ByVal Value As ttStyleEnum)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Style = 5
			mvarStyle = Value
		End Set
	End Property
	
	
	Public Property Centered() As Boolean
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.Centered
			Centered = mvarCentered
		End Get
		Set(ByVal Value As Boolean)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.Centered = 5
			mvarCentered = Value
		End Set
	End Property
	
	
	Public Property Icon() As ttIconType
		Get
			Icon = mvarIcon
		End Get
		Set(ByVal Value As ttIconType)
			mvarIcon = Value
			'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			If m_lTTHwnd <> 0 And Not IsNothing(mvarTitle) And mvarIcon <> ttIconType.TTNoIcon Then
				SendMessage(m_lTTHwnd, TTM_SETTITLE, CInt(mvarIcon), mvarTitle)
			End If
		End Set
	End Property
	
	
	Public Property ForeColor() As Integer
		Get
			ForeColor = mvarForeColor
		End Get
		Set(ByVal Value As Integer)
			mvarForeColor = Value
			If m_lTTHwnd <> 0 Then
				SendMessage(m_lTTHwnd, TTM_SETTIPTEXTCOLOR, mvarForeColor, 0)
			End If
		End Set
	End Property
	
	
	Public Property Title() As String
		Get
			Title = ti.lpStr
		End Get
		Set(ByVal Value As String)
			mvarTitle = Value
			'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			If m_lTTHwnd <> 0 And Not IsNothing(mvarTitle) And mvarIcon <> ttIconType.TTNoIcon Then
				SendMessage(m_lTTHwnd, TTM_SETTITLE, CInt(mvarIcon), mvarTitle)
			End If
		End Set
	End Property
	
	
	Public Property BackColor() As Integer
		Get
			BackColor = mvarBackColor
		End Get
		Set(ByVal Value As Integer)
			mvarBackColor = Value
			If m_lTTHwnd <> 0 Then
				SendMessage(m_lTTHwnd, TTM_SETTIPBKCOLOR, mvarBackColor, 0)
			End If
		End Set
	End Property
	
	
	Public Property TipText() As String
		Get
			TipText = mvarTipText
		End Get
		Set(ByVal Value As String)
			mvarTipText = Value
			ti.lpStr = Value
			If m_lTTHwnd <> 0 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object ti. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                SendMessageInfo(m_lTTHwnd, TTM_UPDATETIPTEXTA, 0, ti)
			End If
		End Set
	End Property
	
	
	Public Property VisibleTime() As Integer
		Get
			VisibleTime = mvarVisibleTime
		End Get
		Set(ByVal Value As Integer)
			mvarVisibleTime = Value
		End Set
	End Property
	
	
	Public Property DelayTime() As Integer
		Get
			DelayTime = mvarDelayTime
		End Get
		Set(ByVal Value As Integer)
			mvarDelayTime = Value
		End Set
	End Property
	
	Public Function Create(ByVal ParentHwnd As Integer, ByRef X As Integer, ByRef Y As Integer) As Boolean
		Dim lWinStyle As Integer
		
		If m_lTTHwnd <> 0 Then
			DestroyWindow(m_lTTHwnd)
		End If
		
		m_lParentHwnd = ParentHwnd
		
		lWinStyle = TTS_ALWAYSTIP Or TTS_NOPREFIX
		
		''create baloon style if desired
		If mvarStyle = ttStyleEnum.TTBalloon Then lWinStyle = lWinStyle Or TTS_BALLOON
		
		
		m_lTTHwnd = CreateWindowEx(0, TOOLTIPS_CLASSA, "SBToolTipWindow", lWinStyle, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, 0, 0, VB6.GetHInstance.ToInt32, 0)
		
		''now set our tooltip info structure
		With ti
			''if we want it centered, then set that flag
			If mvarCentered Then
				.lFlags = TTF_SUBCLASS Or TTF_CENTERTIP Or TTF_IDISHWND
			Else
				.lFlags = TTF_SUBCLASS Or TTF_IDISHWND
			End If
			
			''set the hwnd prop to our parent control's hwnd
			.hWnd = m_lParentHwnd
			.lId = m_lParentHwnd '0
			.hInstance = VB6.GetHInstance.ToInt32
			'.lpstr = ALREADY SET
			'.lpRect = lpRect
			.lSize = Len(ti)
		End With
		
		''add the tooltip structure
		'UPGRADE_WARNING: Couldn't resolve default property of object ti. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        SendMessageInfo(m_lTTHwnd, TTM_ADDTOOLA, 0, ti)
		
		''if we want a title or we want an icon
		If mvarTitle <> vbNullString Or mvarIcon <> ttIconType.TTNoIcon Then
			SendMessage(m_lTTHwnd, TTM_SETTITLE, CInt(mvarIcon), mvarTitle)
		End If
		
		'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If Not IsNothing(mvarForeColor) Then
			SendMessage(m_lTTHwnd, TTM_SETTIPTEXTCOLOR, mvarForeColor, 0)
		End If
		
		'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If Not IsNothing(mvarBackColor) Then
			SendMessage(m_lTTHwnd, TTM_SETTIPBKCOLOR, mvarBackColor, 0)
		End If
		
		SendMessageLong(m_lTTHwnd, TTM_SETDELAYTIME, TTDT_AUTOPOP, mvarVisibleTime)
		SendMessageLong(m_lTTHwnd, TTM_SETDELAYTIME, TTDT_INITIAL, mvarDelayTime)
		
		Create = CBool(m_lTTHwnd)
	End Function
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		InitCommonControls()
		mvarDelayTime = 200
		mvarVisibleTime = 5000
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		Destroy()
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	Public Sub Destroy()
		If m_lTTHwnd <> 0 Then
			DestroyWindow(m_lTTHwnd)
		End If
	End Sub
End Class