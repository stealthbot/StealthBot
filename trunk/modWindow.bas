Attribute VB_Name = "modWindow"
Option Explicit
'modSubclassing - project StealthBot
' authored 7/28/04 andy@stealthbot.net
' updated 4/12/06 to add transparency
' updated 12/24/06 to add hooking for the main send box on frmMain (merry Christmas!)

Private Type NMHDR
    hWndFrom As Long
    idFrom   As Long
    code     As Long
End Type

Private Type CHARRANGE
    cpMin As Long
    cpMax As Long
End Type

Private Type ENLINK
    hdr    As NMHDR
    msg    As Long
    wParam As Long
    lParam As Long
    chrg   As CHARRANGE
End Type

Private Type TEXTRANGE
    chrg      As CHARRANGE
    lpstrText As String
End Type

Public ID_TASKBARICON       As Integer
Public TASKBARCREATED_MSGID As Long

Private Const WM_NOTIFY = &H4E
Private Const EM_SETEVENTMASK = &H445
Private Const EM_GETEVENTMASK = &H43B
Private Const EM_GETTEXTRANGE = &H44B
Private Const EM_AUTOURLDETECT = &H45B
Private Const EN_LINK = &H70B
Private Const CFE_LINK = &H20
Private Const ENM_LINK = &H4000000
Private Const SW_SHOW = 5
Private Const WM_COMMAND = &H111
Private Const WM_USER = &H400
Private Const WM_NCDESTROY = &H82
Public Const WM_ICONNOTIFY = WM_USER + 100

Private hWndSet As New Dictionary
Private hWndRTB As New Dictionary

Public Sub HookWindowProc(ByVal hWnd As Long)

    Dim OldWindowProc As Long
    
    OldWindowProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf NewWindowProc)

    hWndSet(hWnd) = OldWindowProc
 
End Sub

Public Sub UnhookWindowProc(ByVal hWnd As Long)

    SetWindowLong hWnd, GWL_WNDPROC, hWndSet(hWnd)

    hWndSet.Remove hWnd

End Sub

Public Sub EnableURLDetect(ByVal hWndTextbox As Long)

    SendMessage hWndTextbox, EM_SETEVENTMASK, 0, ByVal ENM_LINK Or SendMessage(hWndTextbox, EM_GETEVENTMASK, 0, 0)
    SendMessage hWndTextbox, EM_AUTOURLDETECT, 1, ByVal 0

End Sub

Public Sub DisableURLDetect(ByVal hWndTextbox As Long)

    SendMessage hWndTextbox, EM_AUTOURLDETECT, 0, ByVal 0

End Sub

Public Function NewWindowProc(ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Dim Rezult As Long
    Dim uHead  As NMHDR
    Dim eLink  As ENLINK
    Dim eText  As TEXTRANGE
    Dim sText  As String
    Dim lLen   As Long
    
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
        CopyMemory uHead, ByVal lParam, LenB(uHead)
       
        If (uHead.code = EN_LINK) Then
            CopyMemory eLink, ByVal lParam, LenB(eLink)
       
            With eLink
                If .msg = WM_LBUTTONDBLCLK Then
                    eText.chrg.cpMin = .chrg.cpMin
                    eText.chrg.cpMax = .chrg.cpMax
                    eText.lpstrText = Space$(1024)
       
                    lLen = SendMessageAny(uHead.hWndFrom, EM_GETTEXTRANGE, 0, eText)
                    sText = Left$(eText.lpstrText, lLen)
       
                   ShellExecute hWnd, 0&, sText, 0&, 0&, vbNormalFocus
                End If
            End With
        End If
    ElseIf msg = WM_COMMAND Then
        If lParam = 0 Then
            MenuClick hWnd, wParam
        End If
    End If
    
    NewWindowProc = CallWindowProc(hWndSet(hWnd), hWnd, msg, wParam, lParam)
    
End Function
