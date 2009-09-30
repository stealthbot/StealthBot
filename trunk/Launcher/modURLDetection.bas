Attribute VB_Name = "modURLDetection"
Option Explicit

'Taken from SB with a bit of tweaking

Private Const OBJECT_NAME As String = "modURLDetection"

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, source As Any, ByVal length As Long)
Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

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

Private Const WM_NOTIFY        As Long = &H4E
Private Const WM_LBUTTONDBLCLK As Long = &H203  'Double-click
Private Const EM_GETEVENTMASK  As Long = &H43B
Private Const EM_SETEVENTMASK  As Long = &H445
Private Const EM_GETTEXTRANGE  As Long = &H44B
Private Const EM_AUTOURLDETECT As Long = &H45B
Private Const ENM_LINK         As Long = &H4000000
Private Const EN_LINK          As Long = &H70B
Private Const GWL_WNDPROC      As Long = (-4)
Private Const SW_SHOW          As Long = 5
Private Const WM_MOVE          As Long = &H3

Private m_OldProcs As New Dictionary

Public Sub HookWindowProc(ByRef hWnd As Long)
On Error GoTo ERROR_HANDLER:

    Dim lOldWindowProc As Long
    lOldWindowProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf NewWindowProc)
    m_OldProcs.Add CStr(hWnd), lOldWindowProc
    
    Exit Sub
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "HookWindowProc"
End Sub

Public Sub EnableURLDetection(ByRef hWnd As Long)
On Error GoTo ERROR_HANDLER:
    
    SendMessage hWnd, EM_SETEVENTMASK, 0, ByVal ENM_LINK Or SendMessage(hWnd, EM_GETEVENTMASK, 0, 0)
    SendMessage hWnd, EM_AUTOURLDETECT, 1, ByVal 0
    
    Exit Sub
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "EnableURLDetection"
End Sub

Public Sub UnHookWindowProc(ByRef hWnd As Long)
On Error GoTo ERROR_HANDLER:
    
    If (Not m_OldProcs.Exists(CStr(hWnd))) Then Exit Sub
    
    SetWindowLong hWnd, GWL_WNDPROC, m_OldProcs.Item(CStr(hWnd))
    m_OldProcs.Remove CStr(hWnd)

    Exit Sub
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "UnHookWindowProc"
End Sub

Public Sub UnHookAllProcs()
On Error GoTo ERROR_HANDLER:
    
    Do While m_OldProcs.Count > 0
        UnHookWindowProc CLng(m_OldProcs.Keys(0))
    Loop

    Exit Sub
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "UnHookAllProcs"
End Sub

Public Function NewWindowProc(ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error GoTo ERROR_HANDLER:
    
    Dim mHeader As NMHDR
    Dim eLink   As ENLINK
    Dim eText   As TEXTRANGE
    Dim lRet    As Long

    If (Not m_OldProcs.Exists(CStr(hWnd))) Then Exit Function

    If msg = WM_NOTIFY Then
        CopyMemory mHeader, ByVal lParam, LenB(mHeader)
        
        If (mHeader.code = EN_LINK) Then
            CopyMemory eLink, ByVal lParam, LenB(eLink)
            With eLink
                If (.msg = WM_LBUTTONDBLCLK) Then
                    eText.chrg.cpMax = .chrg.cpMax
                    eText.chrg.cpMin = .chrg.cpMin
                    eText.lpstrText = String$(1025, Chr$(0))
                    
                    lRet = SendMessageAny(mHeader.hWndFrom, EM_GETTEXTRANGE, 0, eText)
                    eText.lpstrText = Left$(eText.lpstrText, lRet)
    
                    ShellExecute hWnd, vbNullString, eText.lpstrText, vbNullString, vbNullString, SW_SHOW
                End If
            End With
        End If
    ElseIf msg = WM_MOVE And hWnd = frmLauncher.hWnd Then
        'If (frmStatus.Visible) Then
            frmStatus.Move frmLauncher.Left + frmLauncher.Width + 100, frmLauncher.Top
        'End If
    End If
    
    NewWindowProc = CallWindowProc(m_OldProcs.Item(CStr(hWnd)), hWnd, msg, wParam, lParam)

    Exit Function
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "NewWindowProc"
End Function

