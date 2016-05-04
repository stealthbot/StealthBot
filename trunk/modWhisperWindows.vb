Option Strict Off
Option Explicit On
Module modWhisperWindows
	
	Public colWhisperWindows As Collection
	
	Public Function AddWhisperWindow(ByVal sUsername As String) As Short
		Dim intRet As Short
		Dim ToAdd As frmWhisperWindow
		
		intRet = WWUserNameToIndex(sUsername)
		
		If intRet = 0 Then
			ToAdd = New frmWhisperWindow
			
			With ToAdd
				.sWhisperTo = sUsername
				.myIndex = colWhisperWindows.Count() + 1
				.StartDate = Now
				.Text = "Whisper Window: " & sUsername
			End With
			
			colWhisperWindows.Add(ToAdd)
			
			intRet = colWhisperWindows.Count()
			ShowWW(intRet)
		End If
		
		AddWhisperWindow = intRet
	End Function
	
	Public Function WWUserNameToIndex(ByVal sUsername As String) As Short
		Dim i As Short
		
		WWUserNameToIndex = 0
		
		If ActiveWWs Then
			For i = 1 To colWhisperWindows.Count()
				'UPGRADE_WARNING: Couldn't resolve default property of object colWhisperWindows.Item().sWhisperTo. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If StrComp(colWhisperWindows.Item(i).sWhisperTo, sUsername, CompareMethod.Text) = 0 Then
					WWUserNameToIndex = i
					Exit For
				End If
			Next i
		End If
	End Function
	
	Public Function ActiveWWs() As Boolean
		On Error Resume Next
		ActiveWWs = (colWhisperWindows.Count() > 0)
	End Function
	
	Public Function ShowWW(ByVal Index As Short) As Boolean
		Dim ReturnFocus As Boolean
		
		ReturnFocus = cboSendHadFocus
		
		If Index > 0 And Index <= colWhisperWindows.Count() Then
			ShowWW = True
			
			'UPGRADE_WARNING: Couldn't resolve default property of object colWhisperWindows.Item(Index).Shown. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Not colWhisperWindows.Item(Index).Shown Then
				'UPGRADE_WARNING: Couldn't resolve default property of object colWhisperWindows.Item().Show. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				colWhisperWindows.Item(Index).Show()
				'UPGRADE_WARNING: Couldn't resolve default property of object colWhisperWindows.Item().Shown. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				colWhisperWindows.Item(Index).Shown = True
				
				If ReturnFocus Then
					frmChat.cboSend.Focus()
				End If
			End If
		Else
			ShowWW = False
		End If
	End Function
	
	Public Function HideWW(ByVal Index As Short) As Boolean
		If Index > 0 And Index <= colWhisperWindows.Count() Then
			HideWW = True
			'UPGRADE_WARNING: Couldn't resolve default property of object colWhisperWindows.Item().Hide. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			colWhisperWindows.Item(Index).Hide()
			'UPGRADE_WARNING: Couldn't resolve default property of object colWhisperWindows.Item().Shown. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			colWhisperWindows.Item(Index).Shown = False
		Else
			HideWW = False
		End If
	End Function
	
	Public Sub HideAllWWs()
		Dim i As Short
		
		If ActiveWWs Then
			For i = 1 To colWhisperWindows.Count()
				HideWW(i)
			Next i
		End If
	End Sub
	
	Public Sub DestroyAllWWs()
		Dim fTemp As frmWhisperWindow
		
		While ActiveWWs
			fTemp = colWhisperWindows.Item(1)
			colWhisperWindows.Remove(1)
			fTemp.Close()
			'UPGRADE_NOTE: Object fTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			fTemp = Nothing
		End While
	End Sub
	
	Public Sub DestroyWW(ByVal Index As Short)
		Dim fTemp As frmWhisperWindow
		If Index > 0 And Index <= colWhisperWindows.Count() Then
			
			fTemp = colWhisperWindows.Item(Index)
			colWhisperWindows.Remove(Index)
			fTemp.Close()
			'UPGRADE_NOTE: Object fTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			fTemp = Nothing
		End If
	End Sub
	
	'Public Function WWNewWndProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
	'    Const WM_ACTIVATE = 6
	'    Const WA_INACTIVE = 0
	'    Dim i As Integer
	'    Dim thisWW As Integer
	'
	'    For i = 1 To colWhisperWindows.Count
	'        If colWhisperWindows.Item(i).hWnd = hWnd Then
	'            thisWW = i
	'            Debug.Print "found ww: " & i
	'            Debug.Print "hwnd: " & hWnd & ", msg: " & Msg & ", wParam: " & wParam & ", lParam: " & lParam
	'            Exit For
	'        End If
	'    Next i
	'
	'    If thisWW > 0 Then
	'        With colWhisperWindows.Item(i)
	'            If Msg = WM_ACTIVATE Then
	'                If ((wParam And &HFFFF) <> WA_INACTIVE) Then
	'                    If (lParam <> 0) Then
	'                        Call SetActiveWindow(lParam)
	'                    Else
	'                        Call SetActiveWindow(0&)
	'                    End If
	'                End If
	'            End If
	'
	'            Debug.Print "Local vars: " & .MyOldWndProc & ", " & .hWnd
	'
	'            WWNewWndProc = CallWindowProc(.MyOldWndProc, .hWnd, Msg, wParam, lParam)
	'        End With
	'    End If
	'End Function
End Module