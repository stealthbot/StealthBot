Option Strict Off
Option Explicit On
Module modMenuObj
	'Using the Menu APIs to Grow or Shrink a Menu During Run-time
	'(c) Jon Vote, 2003
	'
	'Idioma Software Inc.
	'jon@idioma-software.com
	'www.idioma-software.com
	'www.skycoder.com
	
	
	'Adapted to StealthBot
	' 2007-06-10, Andy T
	
	
	Public Declare Function GetMenu Lib "user32" (ByVal hWnd As Integer) As Integer
	Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Integer, ByVal nPos As Integer) As Integer
	Public Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Integer) As Integer
	Public Declare Function CreatePopupMenu Lib "user32" () As Integer
	Public Declare Function GetMenuString Lib "user32"  Alias "GetMenuStringA"(ByVal hMenu As Integer, ByVal wIDItem As Integer, ByVal lpString As String, ByVal nMaxCount As Integer, ByVal wFlag As Integer) As Integer
	Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Integer, ByVal nPos As Integer) As Integer
	Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Integer) As Integer
	'UPGRADE_WARNING: Structure MENUITEMINFO may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Public Declare Function GetMenuItemInfo Lib "user32"  Alias "GetMenuItemInfoA"(ByVal hMenu As Integer, ByVal uItemID As Integer, ByVal ByPosition As Boolean, ByRef lpMenuItemInfo As MENUITEMINFO) As Integer
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Public Declare Function AppendMenu Lib "user32"  Alias "AppendMenuA"(ByVal hMenu As Integer, ByVal wFlags As Integer, ByVal wIDNewItem As Integer, ByVal lpNewItem As Any) As Integer
	Public Declare Function CheckMenuItem Lib "user32" (ByVal hMenu As Integer, ByVal wIDCheckItem As Integer, ByVal wCheck As Integer) As Integer
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Public Declare Function ModifyMenu Lib "user32"  Alias "ModifyMenuA"(ByVal hMenu As Integer, ByVal uPosition As Integer, ByVal uFlags As Integer, ByVal uIDNewItem As Integer, ByVal lpNewItemStr As Any) As Integer
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Public Declare Function InsertMenu Lib "user32"  Alias "InsertMenuA"(ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer, ByVal wIDNewItem As Integer, ByVal lpNewItem As Any) As Integer
	Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer) As Integer
	
	'Possible Values() for wFlags
	Public Const MF_BITMAP As Integer = &H4 'Menu item is bitmap. lpNewItem = handle to bitmap.
	Public Const MF_CHECKED As Integer = &H8 'Check flag.
	Public Const MF_DISABLED As Integer = &H2 'Disable flag.
	Public Const MF_ENABLED As Integer = &H0 'Enable flag.
	Public Const MF_GRAYED As Integer = &H1 'Greyed flag.
	Public Const MF_MENUBARBREAK As Integer = &H20 'Separator - verticle line if popup.
	Public Const MF_MENUBREAK As Integer = &H40 'Separator - no columns.
	Public Const MF_OWNERDRAW As Integer = &H100 'Owner drawn.
	Public Const MF_POPUP As Integer = &H10 'Popup menu (Sub-menu).
	Public Const MF_SEPARATOR As Integer = &H800 'Seperator - dropdown only.
	Public Const MF_STRING As Integer = &H0 'Item is a string.
	Public Const MF_UNCHECKED As Integer = &H0 'Un-check flag.
	
	'Refer to menu item by position or command (ID).
	Public Const MF_BYCOMMAND As Integer = &H0
	Public Const MF_BYPOSITION As Integer = &H400
	
	'Menu Action Enum - possible user responses
	Public Enum MenuAction
		ACTION_CONTINUE = 0
		ACTION_INSERT_ITEM_BEFORE = 1
		ACTION_INSERT_ITEM_AFTER = 2
		ACTION_INSERT_SUBMENU_BEFORE = 3
		ACTION_INSERT_SUBMENU_AFTER = 4
		ACTION_DELETE = 5
	End Enum
	
	Private m_menus As Collection
	
	Public Function DynamicMenus() As Collection
		
		On Error GoTo ERROR_HANDLER
		
		If (m_menus Is Nothing) Then
			m_menus = New Collection
		End If
		
		DynamicMenus = m_menus
		
		Exit Function
		
ERROR_HANDLER: 
		
		frmChat.AddChat(RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & " in modMenuObj.GetMenuByID().")
		
		Resume Next
		
	End Function
	
	Public Function GetMenuByID(ByVal lng As Integer) As Object
		
		On Error GoTo ERROR_HANDLER
		
		Dim i As Short
		
		For i = 1 To DynamicMenus.Count()
			'UPGRADE_WARNING: Couldn't resolve default property of object DynamicMenus(i).ID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If (DynamicMenus.Item(i).ID = lng) Then
				GetMenuByID = DynamicMenus(i)
				
				Exit Function
			End If
		Next i
		
		Exit Function
		
ERROR_HANDLER: 
		
		frmChat.AddChat(RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & " in modMenuObj.GetMenuByID().")
		
		Resume Next
		
	End Function
	
	Public Sub MenuClick(ByRef hWnd As Integer, ByRef lngMenuCommand As Integer)
		
		On Error GoTo ERROR_HANDLER
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByMenuID(lngMenuCommand)
		
		' is this a dynamic scripting menu?
		Dim i As Short
		Dim s_name As String
		Dim sub_name As String
		If (obj.ObjName <> vbNullString) Then
			On Error Resume Next
			
			RunInSingle(obj.SCModule, obj.ObjName & "_Click")
		Else
			
			For i = 1 To DynamicMenus.Count()
				'UPGRADE_WARNING: Couldn't resolve default property of object DynamicMenus(i).ID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (DynamicMenus.Item(i).ID = lngMenuCommand) Then
					' is this a default scripting menu?
					'UPGRADE_WARNING: Couldn't resolve default property of object DynamicMenus().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If (Left(DynamicMenus.Item(i).Name, 1) = Chr(0)) Then
						
						'UPGRADE_WARNING: Couldn't resolve default property of object DynamicMenus().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						s_name = Split(Mid(DynamicMenus.Item(i).Name, 2), Chr(0))(0)
						'UPGRADE_WARNING: Couldn't resolve default property of object DynamicMenus().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						sub_name = Split(Mid(DynamicMenus.Item(i).Name, 2), Chr(0))(1)
						
						If (sub_name = "ENABLE|DISABLE") Then
							'UPGRADE_WARNING: Couldn't resolve default property of object DynamicMenus(i).Checked. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If (DynamicMenus.Item(i).Checked) Then
								ProcessCommand(GetCurrentUsername, "/disable " & s_name, True)
							Else
								ProcessCommand(GetCurrentUsername, "/enable " & s_name, True)
							End If
						ElseIf (sub_name = "VIEW_SCRIPT") Then 
							If (Config.ScriptViewer = vbNullString) Then
								'UPGRADE_WARNING: Couldn't resolve default property of object GetScriptDictionary()(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								ShellOpenURL(GetScriptDictionary(GetModuleByName(s_name))("Path"),  , False)
							Else
								'UPGRADE_WARNING: Couldn't resolve default property of object GetScriptDictionary(GetModuleByName(s_name))(Path). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								Shell(Chr(34) & Config.ScriptViewer & Chr(34) & Space(1) & Chr(34) & GetScriptDictionary(GetModuleByName(s_name))("Path") & Chr(34))
							End If
						End If
					End If
					
					Exit For
				End If
			Next i
		End If
		
		Exit Sub
		
ERROR_HANDLER: 
		
		frmChat.AddChat(RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & " in modMenuObj.MenuClick().")
		
		Resume Next
	End Sub
End Module