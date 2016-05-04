Option Strict Off
Option Explicit On
Module modWindow
	'modSubclassing - project StealthBot
	' authored 7/28/04 andy@stealthbot.net
	' updated 4/12/06 to add transparency
	' updated 12/24/06 to add hooking for the main send box on frmMain (merry Christmas!)
	
	Private Structure NMHDR
		Dim hWndFrom As Integer
		Dim idFrom As Integer
		Dim code As Integer
	End Structure
	
	Private Structure CHARRANGE
		Dim cpMin As Integer
		Dim cpMax As Integer
	End Structure
	
	Private Structure ENLINK
		Dim hdr As NMHDR
		Dim Msg As Integer
		Dim wParam As Integer
		Dim lParam As Integer
		Dim chrg As CHARRANGE
	End Structure
	
	Private Structure TEXTRANGE
		Dim chrg As CHARRANGE
		Dim lpstrText As String
	End Structure
	
	Private Structure COPYDATASTRUCT
		Dim dwData As Integer
		Dim cbData As Integer
		Dim lpData As Integer
	End Structure
	
	Public ID_TASKBARICON As Short
	Public TASKBARCREATED_MSGID As Integer
	
	' windows messages
	Private Const WM_NOTIFY As Integer = &H4E
	Private Const WM_COMMAND As Integer = &H111
	Private Const WM_USER As Integer = &H400
	Private Const WM_NCDESTROY As Integer = &H82
	Private Const WM_COPYDATA As Integer = &H4A
	Public Const WM_ICONNOTIFY As Decimal = WM_USER + 100
	' RTB rich edit control messages
	Private Const EM_SETEVENTMASK As Integer = &H445
	Private Const EM_GETEVENTMASK As Integer = &H43B
	Private Const EM_GETTEXTRANGE As Integer = &H44B
	Private Const EM_AUTOURLDETECT As Integer = &H45B
	' RTB rich edit notifications
	Private Const EN_LINK As Integer = &H70B
	' EN_LINK effects
	Private Const CFE_LINK As Integer = &H20
	' EN_LINK message flag
	Private Const ENM_LINK As Integer = &H4000000
	' show window function
	Private Const SW_SHOW As Short = 5
	' list view notifications
	Private Const LVN_FIRST As Short = -100
	Private Const LVN_BEGINDRAG As Short = (LVN_FIRST - 9)
	
	Private hWndSet As New Scripting.Dictionary
	Private hWndRTB As New Scripting.Dictionary
	
	Public Sub HookWindowProc(ByVal hWnd As Integer)
		
		Dim OldWindowProc As Integer
		
		'UPGRADE_WARNING: Add a delegate for AddressOf NewWindowProc Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="E9E157F7-EF0C-4016-87B7-7D7FBBC6EE08"'
		OldWindowProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf NewWindowProc)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object hWndSet(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		hWndSet(hWnd) = OldWindowProc
		
	End Sub
	
	Public Sub UnhookWindowProc(ByVal hWnd As Integer)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object hWndSet(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		SetWindowLong(hWnd, GWL_WNDPROC, hWndSet(hWnd))
		
		hWndSet.Remove(hWnd)
		
	End Sub
	
	Public Sub EnableURLDetect(ByVal hWndTextbox As Integer)
		
		SendMessage(hWndTextbox, EM_SETEVENTMASK, 0, ENM_LINK Or SendMessage(hWndTextbox, EM_GETEVENTMASK, 0, 0))
		SendMessage(hWndTextbox, EM_AUTOURLDETECT, 1, 0)
		
	End Sub
	
	Public Sub DisableURLDetect(ByVal hWndTextbox As Integer)
		
		SendMessage(hWndTextbox, EM_AUTOURLDETECT, 0, 0)
		
	End Sub
	
	Public Function NewWindowProc(ByVal hWnd As Integer, ByVal Msg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
		
		Dim Rezult As Integer
		Dim uHead As NMHDR
		Dim eLink As ENLINK
		Dim eText As TEXTRANGE
		Dim sText As String
		Dim lLen As Integer
		Dim cds As COPYDATASTRUCT
		Dim buf(255) As Byte
		Dim Data As String
		
		If Msg = TASKBARCREATED_MSGID Then
			Shell_NotifyIcon(NIM_ADD, nid)
		End If
		
		If wParam = ID_TASKBARICON Then
			Select Case lParam
				Case WM_LBUTTONUP
					frmChat.WindowState = System.Windows.Forms.FormWindowState.Normal
					Rezult = SetForegroundWindow(frmChat.Handle.ToInt32)
					frmChat.Show()
				Case WM_RBUTTONUP
					SetForegroundWindow(frmChat.Handle.ToInt32)
					'UPGRADE_ISSUE: Form method frmChat.PopupMenu was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					frmChat.PopupMenu(frmChat.mnuTray)
			End Select
		End If
		
		If Msg = WM_NOTIFY Then
            'UPGRADE_WARNING: Couldn't resolve default property of object uHead. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            CopyMemory(uHead, lParam, Len(uHead))
			
			If (uHead.code = EN_LINK) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object eLink. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                CopyMemory(eLink, lParam, Len(eLink))
				
				With eLink
					If .Msg = WM_LBUTTONDBLCLK Then
						eText.chrg.cpMin = .chrg.cpMin
						eText.chrg.cpMax = .chrg.cpMax
						eText.lpstrText = Space(1024)
						
						'UPGRADE_WARNING: Couldn't resolve default property of object eText. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						lLen = SendMessageAny(uHead.hWndFrom, EM_GETTEXTRANGE, 0, eText)
						sText = Left(eText.lpstrText, lLen)
						
						ShellOpenURL(sText,  , False)
					End If
				End With
				
				' See if this is the start of a drag.
				'UPGRADE_WARNING: Couldn't resolve default property of object LVN_BEGINDRAG. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf uHead.code = LVN_BEGINDRAG Then 
				' A drag is beginning. Ignore this event.
				' Indicate we have handled this.
				NewWindowProc = 1
				' Do nothing else.
				Exit Function
			End If
		ElseIf Msg = WM_COMMAND Then 
			If lParam = 0 Then
				MenuClick(hWnd, wParam)
			End If
		ElseIf Msg = WM_COPYDATA Then 
			'UPGRADE_WARNING: Couldn't resolve default property of object cds. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Call CopyMemory(cds, lParam, Len(cds))
			If (cds.cbData < UBound(buf)) Then
				Call CopyMemory(buf(0), cds.lpData, cds.cbData)
				'UPGRADE_ISSUE: Constant vbUnicode was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
				Data = StrConv(System.Text.UnicodeEncoding.Unicode.GetString(buf), vbUnicode)
				Data = Left(Data, InStr(1, Data, Chr(0)) - 1)
				If (StrComp(Data, "-reloadscripts", CompareMethod.Text) = 0) Then
					SharedScriptSupport.ReloadScript()
				End If
			End If
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object hWndSet(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		NewWindowProc = CallWindowProc(hWndSet(hWnd), hWnd, Msg, wParam, lParam)
		
	End Function
End Module