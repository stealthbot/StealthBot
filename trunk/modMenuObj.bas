Attribute VB_Name = "modMenuObj"
'Using the Menu APIs to Grow or Shrink a Menu During Run-time
'(c) Jon Vote, 2003
'
'Idioma Software Inc.
'jon@idioma-software.com
'www.idioma-software.com
'www.skycoder.com


'Adapted to StealthBot
' 2007-06-10, Andy T

Option Explicit

Public Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function CreatePopupMenu Lib "user32" () As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal uItemID As Long, ByVal ByPosition As Boolean, lpMenuItemInfo As MENUITEMINFO) As Long
Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Public Declare Function CheckMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDCheckItem As Long, ByVal wCheck As Long) As Long
Public Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal uPosition As Long, ByVal uFlags As Long, ByVal uIDNewItem As Long, ByVal lpNewItemStr As Any) As Long
Public Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long

'Possible Values() for wFlags
Public Const MF_BITMAP = &H4&        'Menu item is bitmap. lpNewItem = handle to bitmap.
Public Const MF_CHECKED = &H8&       'Check flag.
Public Const MF_DISABLED = &H2&      'Disable flag.
Public Const MF_ENABLED = &H0&       'Enable flag.
Public Const MF_GRAYED = &H1&        'Greyed flag.
Public Const MF_MENUBARBREAK = &H20& 'Separator - verticle line if popup.
Public Const MF_MENUBREAK = &H40&    'Separator - no columns.
Public Const MF_OWNERDRAW = &H100&   'Owner drawn.
Public Const MF_POPUP = &H10&        'Popup menu (Sub-menu).
Public Const MF_SEPARATOR = &H800&   'Seperator - dropdown only.
Public Const MF_STRING = &H0&        'Item is a string.
Public Const MF_UNCHECKED = &H0&     'Un-check flag.
 
'Refer to menu item by position or command (ID).
Public Const MF_BYCOMMAND = &H0&
Public Const MF_BYPOSITION = &H400&

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
        Set m_menus = New Collection
    End If
    
    Set DynamicMenus = m_menus
    
    Exit Function
    
ERROR_HANDLER:
    
    frmChat.AddChat RTBColors.ErrorMessageText, _
        "Error (#" & Err.Number & "): " & Err.description & " in modMenuObj.GetMenuByID()."
        
    Resume Next

End Function

Public Function GetMenuByID(ByVal lng As Long) As Object

    On Error GoTo ERROR_HANDLER

    Dim i As Integer
    
    For i = 1 To DynamicMenus.Count
        If (DynamicMenus(i).ID = lng) Then
            Set GetMenuByID = DynamicMenus(i)
            
            Exit Function
        End If
    Next i

    Exit Function
    
ERROR_HANDLER:
    
    frmChat.AddChat RTBColors.ErrorMessageText, _
        "Error (#" & Err.Number & "): " & Err.description & " in modMenuObj.GetMenuByID()."
        
    Resume Next

End Function

Public Sub MenuClick(hWnd As Long, lngMenuCommand As Long)
    
    On Error GoTo ERROR_HANDLER
    
    Dim obj As scObj 

    obj = GetScriptObjByMenuID(lngMenuCommand)
    
    ' is this a dynamic scripting menu?
    If (obj.ObjName <> vbNullString) Then
        On Error Resume Next

        RunInSingle obj.SCModule, obj.ObjName & "_Click"
    Else
        Dim i As Integer 
        
        For i = 1 To DynamicMenus.Count
            If (DynamicMenus(i).ID = lngMenuCommand) Then
                ' is this a default scripting menu?
                If (Left$(DynamicMenus(i).Name, 1) = Chr$(0)) Then
                    Dim s_name   As String
                    Dim sub_name As String 

                    s_name = _
                        Split(Mid$(DynamicMenus(i).Name, 2), Chr$(0))(0)
                    sub_name = _
                        Split(Mid$(DynamicMenus(i).Name, 2), Chr$(0))(1)
                        
                    If (sub_name = "ENABLE|DISABLE") Then
                        If (DynamicMenus(i).Checked) Then
                            ProcessCommand GetCurrentUsername, "/disable " & s_name, True
                        Else
                            ProcessCommand GetCurrentUsername, "/enable " & s_name, True
                        End If
                    ElseIf (sub_name = "VIEW_SCRIPT") Then
                        If (ReadCfg("Override", "ScriptViewer") = vbNullString) Then
                            ShellExecute frmChat.hWnd, "Open", GetScriptDictionary(GetModuleByName(s_name))("Path"), vbNullString, vbNullString, _
                                vbNormalFocus
                        Else
                            Shell Chr(34) & ReadCfg("Override", "ScriptViewer") & Chr(34) & Space(1) & Chr(34) & GetScriptDictionary(GetModuleByName(s_name))("Path") & Chr(34)
                        End If
                    End If
                End If
                
                Exit For
            End If
        Next i
    End If
    
    Exit Sub
    
ERROR_HANDLER:
    
    frmChat.AddChat RTBColors.ErrorMessageText, _
        "Error (#" & Err.Number & "): " & Err.description & " in modMenuObj.MenuClick()."
        
    Resume Next
End Sub
