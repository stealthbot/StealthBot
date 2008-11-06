Attribute VB_Name = "modSubclassing"
Option Explicit
'modSubclassing - project StealthBot
' authored 7/28/04 andy@stealthbot.net
' updated 4/12/06 to add transparency
' updated 12/24/06 to add hooking for the main send box on frmMain (merry Christmas!)

Public OldWindowProc As Long
Public hWndSet As Long

Public SendBox_OldWindowProc As Long
Public SendBox_hWndSet As Long

Private Const WM_COMMAND = &H111                     'Used in SendMessage call
Private Const WM_USER = &H400
Public Const WM_NCDESTROY = &H82
Public ID_TASKBARICON As Integer
Public Const WM_ICONNOTIFY = WM_USER + 100
Public TASKBARCREATED_MSGID As Long

Public Sub HookWindowProc(ByVal hWnd As Long)
    If OldWindowProc = 0 Then
        OldWindowProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf NewWindowProc)
        hWndSet = hWnd
    End If
End Sub

Public Sub UnhookWindowProc()
    If OldWindowProc > 0 Then
        OldWindowProc = SetWindowLong(hWndSet, GWL_WNDPROC, OldWindowProc)
        OldWindowProc = 0
    End If
End Sub

Public Sub HookSendBoxWindowProc(ByVal hWnd As Long)
    If SendBox_OldWindowProc = 0 Then
        SendBox_OldWindowProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf NewSendBoxWindowProc)
        SendBox_hWndSet = hWnd
    End If
End Sub

Public Sub UnhookSendBoxWindowProc()
    If SendBox_OldWindowProc > 0 Then
        SendBox_OldWindowProc = SetWindowLong(SendBox_hWndSet, GWL_WNDPROC, SendBox_OldWindowProc)
        SendBox_OldWindowProc = 0
    End If
End Sub

Public Function NewWindowProc(ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error GoTo ERROR_HANDLER
    
    Dim Rezult As Long
    Dim uHead As NMHDR
    Dim eLink As ENLINK
    Dim eText As TEXTRANGE
    Dim sText As String
    Dim lLen As Long
    
    If msg = TASKBARCREATED_MSGID Then
        Shell_NotifyIcon NIM_ADD, nid
    End If
    
    If wParam = ID_TASKBARICON Then
        Select Case lParam
            Case WM_LBUTTONUP
                frmChat.WindowState = vbNormal
                Rezult = SetForegroundWindow(frmChat.hWnd)
                frmChat.Show
            Case WM_RBUTTONUP
                SetForegroundWindow frmChat.hWnd
                frmChat.PopupMenu frmChat.mnuTray
        End Select
    End If
    
    If msg = WM_NOTIFY Then
        CopyMemory uHead, ByVal lParam, Len(uHead)
        
        If (uHead.hWndFrom = hWndRTB) And (uHead.code = EN_LINK) Then
            CopyMemory eLink, ByVal lParam, Len(eLink)
            
            With eLink
                If .msg = WM_LBUTTONDBLCLK Then
                    eText.chrg.cpMin = .chrg.cpMin
                    eText.chrg.cpMax = .chrg.cpMax
                    eText.lpstrText = Space$(1024)
                    
                    lLen = SendMessageAny(hWndRTB, EM_GETTEXTRANGE, 0, eText)
                    sText = Left$(eText.lpstrText, lLen)
                    
                    ShellExecute hWnd, vbNullString, sText, vbNullString, vbNullString, SW_SHOW
                End If
            End With
        End If
    ElseIf msg = WM_COMMAND Then
        If lParam = 0 Then
            Call ProcessMenu(hWnd, wParam)
        End If
    End If
    
    NewWindowProc = CallWindowProc(OldWindowProc, hWndSet, msg, wParam, lParam)
    
    Exit Function
    
ERROR_HANDLER:
    Call frmChat.AddChat(vbRed, "Error: " & Err.description & " in NewWindowProc().")
    
    Exit Function
End Function

Public Function NewSendBoxWindowProc(ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Const WM_ACTIVATE As Long = 6
    
    NewSendBoxWindowProc = CallWindowProc(SendBox_OldWindowProc, SendBox_hWndSet, msg, wParam, lParam)
End Function

Public Sub SetTransparency()
    
End Sub
