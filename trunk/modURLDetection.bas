Attribute VB_Name = "modURLDetection"
Option Explicit

'This module thanks to LordNevar

Public Type NMHDR
    hWndFrom As Long
    idFrom As Long
    code As Long
End Type

Public Type CHARRANGE
    cpMin As Long
    cpMax As Long
End Type

Public Type ENLINK
    hdr As NMHDR
    Msg As Long
    wParam As Long
    lParam As Long
    chrg As CHARRANGE
End Type

Public Type TEXTRANGE
    chrg As CHARRANGE
    lpstrText As String
End Type

Public Const WM_NOTIFY = &H4E
Public Const EM_SETEVENTMASK = &H445
Public Const EM_GETEVENTMASK = &H43B
Public Const EM_GETTEXTRANGE = &H44B
Public Const EM_AUTOURLDETECT = &H45B
Public Const EN_LINK = &H70B

Public Const CFE_LINK = &H20
Public Const ENM_LINK = &H4000000
Public Const SW_SHOW = 5

Public hWndRTB As Long
Private Enabled As Boolean

Public Sub EnableURLDetect(ByVal hWndTextbox As Long)
    If Not Enabled Then
        Enabled = True
        SendMessage hWndTextbox, EM_SETEVENTMASK, 0, ByVal ENM_LINK Or SendMessage(hWndTextbox, EM_GETEVENTMASK, 0, 0)
        SendMessage hWndTextbox, EM_AUTOURLDETECT, 1, ByVal 0
        hWndRTB = hWndTextbox
    End If
End Sub

Public Sub DisableURLDetect(ByVal hWndTextbox As Long)
    If Enabled Then
        Enabled = False
        SendMessage hWndRTB, EM_AUTOURLDETECT, 0, ByVal 0
    End If
End Sub
