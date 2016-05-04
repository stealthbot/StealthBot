Option Strict Off
Option Explicit On
Module modSystray
	
	'UPGRADE_WARNING: Structure NOTIFYICONDATA may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Public Declare Function Shell_NotifyIcon Lib "shell32.dll"  Alias "Shell_NotifyIconA"(ByVal dwMessage As Integer, ByRef lpData As NOTIFYICONDATA) As Integer
	
	Public Structure NOTIFYICONDATA
		Dim cbSize As Integer
		Dim hWnd As Integer
		Dim uId As Integer
		Dim uFlags As Integer
		Dim uCallBackMessage As Integer
		Dim hIcon As Integer
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(64),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=64)> Public szTip() As Char
	End Structure
	
	Public Structure MENUITEMINFO
		Dim cbSize As Integer
		Dim fMask As Integer
		Dim fType As Integer
		Dim fState As Integer
		Dim wID As Integer
		Dim hSubMenu As Integer
		Dim hbmpChecked As Integer
		Dim hbmpUnchecked As Integer
		Dim dwItemData As Integer
		Dim dwTypeData As String
		Dim cch As Integer
	End Structure
	
	'constants required by Shell_NotifyIcon API call:
	Public Const NIM_ADD As Integer = &H0
	Public Const NIM_MODIFY As Integer = &H1
	Public Const NIM_DELETE As Integer = &H2
	Public Const NIF_MESSAGE As Integer = &H1
	Public Const NIF_ICON As Integer = &H2
	Public Const NIF_TIP As Integer = &H4
	Public Const WM_MOUSEMOVE As Integer = &H200
	Public Const WM_LBUTTONDOWN As Integer = &H201 'Button down
	Public Const WM_LBUTTONUP As Integer = &H202 'Button up
	Public Const WM_LBUTTONDBLCLK As Integer = &H203 'Double-click
	Public Const WM_RBUTTONDOWN As Integer = &H204 'Button down
	Public Const WM_RBUTTONUP As Integer = &H205 'Button up
	Public Const WM_RBUTTONDBLCLK As Integer = &H206 'Double-click
	
	Public nid As NOTIFYICONDATA
End Module