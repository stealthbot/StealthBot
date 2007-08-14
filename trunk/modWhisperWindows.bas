Attribute VB_Name = "modWhisperWindows"
Option Explicit

Public colWhisperWindows As Collection

Public Function AddWhisperWindow(ByVal sUsername As String) As Integer
    Dim intRet As Integer
    Dim ToAdd As frmWhisperWindow
    
    intRet = WWUserNameToIndex(sUsername)
    
    If intRet = 0 Then
        Set ToAdd = New frmWhisperWindow
        
        With ToAdd
            .sWhisperTo = sUsername
            .myIndex = colWhisperWindows.Count + 1
            .StartDate = Now
            .Caption = "Whisper Window: " & sUsername
        End With
        
        colWhisperWindows.Add ToAdd
        
        intRet = colWhisperWindows.Count
        ShowWW intRet
    End If
    
    AddWhisperWindow = intRet
End Function

Public Function WWUserNameToIndex(ByVal sUsername As String) As Integer
    Dim i As Integer
    
    WWUserNameToIndex = 0
    
    If ActiveWWs Then
        For i = 1 To colWhisperWindows.Count
            If StrComp(colWhisperWindows.Item(i).sWhisperTo, sUsername, vbTextCompare) = 0 Then
                WWUserNameToIndex = i
                Exit For
            End If
        Next i
    End If
End Function

Public Function ActiveWWs() As Boolean
    On Error Resume Next
    ActiveWWs = (colWhisperWindows.Count > 0)
End Function

Public Function ShowWW(ByVal Index As Integer) As Boolean
    Dim ReturnFocus As Boolean
    
    ReturnFocus = cboSendHadFocus
    
    If Index > 0 And Index <= colWhisperWindows.Count Then
        ShowWW = True
        
        If Not colWhisperWindows.Item(Index).Shown Then
            colWhisperWindows.Item(Index).Show
            colWhisperWindows.Item(Index).Shown = True
            
            If ReturnFocus Then
                frmChat.cboSend.SetFocus
            End If
        End If
    Else
        ShowWW = False
    End If
End Function

Public Function HideWW(ByVal Index As Integer) As Boolean
    If Index > 0 And Index <= colWhisperWindows.Count Then
        HideWW = True
        colWhisperWindows.Item(Index).Hide
        colWhisperWindows.Item(Index).Shown = False
    Else
        HideWW = False
    End If
End Function

Public Sub HideAllWWs()
    Dim i As Integer
    
    If ActiveWWs Then
        For i = 1 To colWhisperWindows.Count
            HideWW (i)
        Next i
    End If
End Sub

Public Sub DestroyAllWWs()
    Dim fTemp As frmWhisperWindow
    
    While ActiveWWs
        Set fTemp = colWhisperWindows.Item(1)
        colWhisperWindows.Remove 1
        Unload fTemp
        Set fTemp = Nothing
    Wend
End Sub

Public Sub DestroyWW(ByVal Index As Integer)
    If Index > 0 And Index <= colWhisperWindows.Count Then
        Dim fTemp As frmWhisperWindow
        
        Set fTemp = colWhisperWindows.Item(Index)
        colWhisperWindows.Remove Index
        Unload fTemp
        Set fTemp = Nothing
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
